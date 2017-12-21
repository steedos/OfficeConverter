﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Win32;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;
using OpenMcdf;
using ExcelInterop = Microsoft.Office.Interop.Excel;

/*
   Copyright 2014-2015 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

namespace OfficeConverter
{
    /// <summary>
    /// This class is used as a placeholder for al Excel related methods
    /// </summary>
    internal static class Excel
    {
        #region Private class ShapePosition
        /// <summary>
        /// Placeholder for shape information
        /// </summary>
        private class ShapePosition
        {
            public int TopLeftColumn { get; private set; }
            public int TopLeftRow { get; private set; }
            public int BottomRightColumn { get; private set; }
            public int BottomRightRow { get; private set; }

            public ShapePosition(ExcelInterop.Shape shape)
            {
                var topLeftCell = shape.TopLeftCell;
                var bottomRightCell = shape.BottomRightCell;
                TopLeftRow = topLeftCell.Row;
                TopLeftColumn = topLeftCell.Column;
                BottomRightRow = bottomRightCell.Row;
                BottomRightColumn = bottomRightCell.Column;
                Marshal.ReleaseComObject(topLeftCell);
                Marshal.ReleaseComObject(bottomRightCell);
            }
        }
        #endregion

        #region Private class ExcelPaperSize
        /// <summary>
        /// Placeholder for papersize and orientation information
        /// </summary>
        private class ExcelPaperSize
        {
            public ExcelInterop.XlPaperSize PaperSize { get; private set; }
            public ExcelInterop.XlPageOrientation Orientation { get; private set; }

            public ExcelPaperSize(ExcelInterop.XlPaperSize paperSize, ExcelInterop.XlPageOrientation orientation)
            {
                PaperSize = paperSize;
                Orientation = orientation;
            }
        }
        #endregion

        #region Private enum MergedCellSearchOrder
        /// <summary>
        /// Direction to search in merged cell
        /// </summary>
        private enum MergedCellSearchOrder
        {
            /// <summary>
            /// Search for first row in the merge area
            /// </summary>
            FirstRow,

            /// <summary>
            /// Search for first column in the merge area
            /// </summary>
            FirstColumn,

            /// <summary>
            /// Search for last row in the merge area
            /// </summary>
            LastRow,

            /// <summary>
            /// Search for last column in the merge area
            /// </summary>
            LastColumn
        }
        #endregion

        #region Fields
        /// <summary>
        /// Excel version number
        /// </summary>
        private static readonly int VersionNumber;

        /// <summary>
        /// Excel maximum rows
        /// </summary>
        private static readonly int MaxRows;

        /// <summary>
        /// Paper sizes to use when detecting optimal page size with the <see cref="SetWorkSheetPaperSize"/> method
        /// </summary>
        private static readonly List<ExcelPaperSize> PaperSizes = new List<ExcelPaperSize>
        {
            new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA4, ExcelInterop.XlPageOrientation.xlPortrait),
            new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA4, ExcelInterop.XlPageOrientation.xlLandscape),
            new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA3, ExcelInterop.XlPageOrientation.xlLandscape),
            new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA3, ExcelInterop.XlPageOrientation.xlPortrait)
        };

        /// <summary>
        /// Zoom ration to use when detecting optimal page size with the <see cref="SetWorkSheetPaperSize"/> method
        /// </summary>
        private static readonly List<int> ZoomRatios = new List<int> { 100, 95, 90, 85, 80, 75 };     
        #endregion

        #region Constructor
        /// <summary>
        /// This constructor is called the first time when the <see cref="Convert"/> or
        /// <see cref="IsPasswordProtected"/> method is called. Some checks are done to
        /// see if all requirements for a succesfull conversion are there.
        /// </summary>
        /// <exception cref="OCConfiguration">Raised when the registry could not be read to determine Excel version</exception>
        static Excel()
        {
            try
            {
                var baseKey = Registry.ClassesRoot;
                var subKey = baseKey.OpenSubKey(@"Excel.Application\CurVer");
                if (subKey != null)
                {
                    switch (subKey.GetValue(string.Empty).ToString().ToUpperInvariant())
                    {
                        // Excel 2003
                        case "EXCEL.APPLICATION.11":
                            VersionNumber = 11;
                            break;

                        // Excel 2007
                        case "EXCEL.APPLICATION.12":
                            VersionNumber = 12;
                            break;

                        // Excel 2010
                        case "EXCEL.APPLICATION.14":
                            VersionNumber = 14;
                            break;

                        // Excel 2013
                        case "EXCEL.APPLICATION.15":
                            VersionNumber = 15;
                            break;

                        // Excel 2016
                        case "EXCEL.APPLICATION.16":
                            VersionNumber = 16;
                            break;

                        default:
                            throw new OCConfiguration("Could not determine Excel version");
                    }
                }
                else
                    throw new OCConfiguration("Could not find registry key Excel.Application\\CurVer");
            }
            catch (Exception exception)
            {
                throw new OCConfiguration("Could not read registry to check Excel version", exception);
            }

            const int excelMaxRowsFrom2003AndBelow = 65535;
            const int excelMaxRowsFrom2007AndUp = 1048576;

            switch (VersionNumber)
            {
                // Excel 2007
                case 12:
                // Excel 2010
                case 14:
                // Excel 2013
                case 15:
                //Excel 2016
                case 16:
                    MaxRows = excelMaxRowsFrom2007AndUp;
                    break;

                // Excel 2003 and older
                default:
                    MaxRows = excelMaxRowsFrom2003AndBelow;
                    break;
            }

            CheckIfSystemProfileDesktopDirectoryExists();
            CheckIfPrinterIsInstalled();
        }
        #endregion

        #region CheckIfSystemProfileDesktopDirectoryExists
        /// <summary>
        /// If you want to run this code on a server then the following folders must exist, if they don't
        /// then you can't use Excel to convert files to PDF
        /// </summary>
        /// <exception cref="OCConfiguration">Raised when the needed directory could not be created</exception>
        private static void CheckIfSystemProfileDesktopDirectoryExists()
        {
            if (Environment.Is64BitOperatingSystem)
            {
                var x64DesktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    @"SysWOW64\config\systemprofile\desktop");

                if (!Directory.Exists(x64DesktopPath))
                {
                    try
                    {
                        Directory.CreateDirectory(x64DesktopPath);
                    }
                    catch (Exception exception)
                    {
                        throw new OCConfiguration("Can't create folder '" + x64DesktopPath +
                                                  "' Excel needs this folder to work on a server, error: " +
                                                  ExceptionHelpers.GetInnerException(exception));
                    }
                }
            }

            var x86DesktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                @"System32\config\systemprofile\desktop");

            if (!Directory.Exists(x86DesktopPath))
            {
                try
                {
                    Directory.CreateDirectory(x86DesktopPath);
                }
                catch (Exception exception)
                {
                    throw new OCConfiguration("Can't create folder '" + x86DesktopPath +
                                              "' Excel needs this folder to work on a server, error: " +
                                              ExceptionHelpers.GetInnerException(exception));
                }
            }
        }
        #endregion

        #region CheckIfPrinterIsInstalled
        /// <summary>
        /// Excel needs a default printer to export to PDF, this method will check if there is one
        /// </summary>
        /// <exception cref="OCConfiguration">Raised when an default printer does not exists</exception>
        private static void CheckIfPrinterIsInstalled()
        {
            var result = false;

            foreach (string printerName in PrinterSettings.InstalledPrinters)
            {
                // Retrieve the printer settings.
                var printer = new PrinterSettings { PrinterName = printerName };

                // Check that this is a valid printer.
                // (This step might be required if you read the printer name
                // from a user-supplied value or a registry or configuration file
                // setting.)
                if (printer.IsValid)
                {
                    result = true;
                    break;
                }
            }

            if (!result)
                throw new OCConfiguration("There is no default printer installed, Excel needs one to export to PDF");
        }
        #endregion

        #region GetColumnAddress
        /// <summary>
        /// Returns the column address for the given <paramref name="column"/>
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        private static string GetColumnAddress(int column)
        {
            if (column <= 26)
                return System.Convert.ToChar(column + 64).ToString(CultureInfo.InvariantCulture);

            var div = column / 26;
            var mod = column % 26;
            if (mod != 0) return GetColumnAddress(div) + GetColumnAddress(mod);
            mod = 26;
            div--;

            return GetColumnAddress(div) + GetColumnAddress(mod);
        }
        #endregion

        #region GetColumnNumber
        /// <summary>
        /// Returns the column number for the given <paramref name="columnAddress"/>
        /// </summary>
        /// <param name="columnAddress"></param>
        /// <returns></returns>
        // ReSharper disable once UnusedMember.Local
        private static int GetColumnNumber(string columnAddress)
        {
            var digits = new int[columnAddress.Length];

            for (var i = 0; i < columnAddress.Length; ++i)
                digits[i] = System.Convert.ToInt32(columnAddress[i]) - 64;

            var mul = 1;
            var res = 0;

            for (var pos = digits.Length - 1; pos >= 0; --pos)
            {
                res += digits[pos] * mul;
                mul *= 26;
            }

            return res;
        }
        #endregion

        #region CheckForMergedCell
        /// <summary>
        /// Checks if the given cell is merged and if so returns the last column or row from this merge.
        /// When the cell is not merged it just returns the cell
        /// </summary>
        /// <param name="range">The cell</param>
        /// <param name="searchOrder"><see cref="MergedCellSearchOrder"/></param>
        /// <returns></returns>
        private static int CheckForMergedCell(ExcelInterop.Range range, MergedCellSearchOrder searchOrder)
        {
            if (range == null)
                return 0;

            var result = 0;
            var mergeArea = range.MergeArea;

            switch (searchOrder)
            {
                case MergedCellSearchOrder.FirstRow:
                    result = mergeArea.Row;
                    break;

                case MergedCellSearchOrder.FirstColumn:
                    result = mergeArea.Column;
                    break;

                case MergedCellSearchOrder.LastRow:
                {
                    result = range.Row;
                    var entireRow = range.EntireRow;

                    for (var i = 1; i < range.Column; i++)
                    {
                        var cell = (ExcelInterop.Range) entireRow.Cells[i];
                        var cellMergeArea = cell.MergeArea;
                        var cellMergeAreaRows = cellMergeArea.Rows;
                        var count = cellMergeAreaRows.Count;

                        Marshal.ReleaseComObject(cellMergeAreaRows);
                        Marshal.ReleaseComObject(cellMergeArea);
                        Marshal.ReleaseComObject(cell);

                        var tempResult = result;

                        if (count > 1 && range.Row + count > tempResult)
                            tempResult = result + count;

                        result = tempResult;
                    }

                    Marshal.ReleaseComObject(entireRow);

                    break;
                }

                case MergedCellSearchOrder.LastColumn:
                {
                    result = range.Column;
                    var columns = mergeArea.Columns;
                    var count = columns.Count;

                    if (count > 1)
                        result += count;

                    Marshal.ReleaseComObject(columns);

                    break;
                }
            }

            if (mergeArea != null)
                Marshal.ReleaseComObject(mergeArea);

            return result;
        }
        #endregion

        #region GetWorksheetPrintArea
        /// <summary>
        /// Figures out the used cell range. This are the cell's that contain any form of text and 
        /// returns this range. An empty range will be returned when there are shapes used on a worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        private static string GetWorksheetPrintArea(ExcelInterop._Worksheet worksheet)
        {
            var firstColumn = 1;
            var firstRow = 1;

            var shapesPosition = new List<ShapePosition>();

            // We can't use this method when there are shapes on a sheet so
            // we return an empty string
            var shapes = worksheet.Shapes;
            if (shapes.Count > 0)
            {
                if (VersionNumber < 14)
                    return "shapes";

                // The shape TopLeftCell and BottomRightCell is only supported from Excel 2010 and up
                foreach (ExcelInterop.Shape shape in worksheet.Shapes)
                {
                    if (shape.AutoShapeType != MsoAutoShapeType.msoShapeMixed)
                        shapesPosition.Add(new ShapePosition(shape));

                    Marshal.ReleaseComObject(shape);
                }

                Marshal.ReleaseComObject(shapes);
            }

            var range = worksheet.Cells[1, 1] as ExcelInterop.Range;
            if (range == null || range.Value == null)
            {
                if (range != null)
                    Marshal.ReleaseComObject(range);

                var firstCellByColumn = worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByColumns);
                var foundByFirstColumn = false;
                if (firstCellByColumn != null)
                {
                    foundByFirstColumn = true;
                    firstColumn = CheckForMergedCell(firstCellByColumn, MergedCellSearchOrder.FirstColumn);
                    firstRow = CheckForMergedCell(firstCellByColumn, MergedCellSearchOrder.FirstRow);
                    Marshal.ReleaseComObject(firstCellByColumn);
                }

                // Search the first used cell row wise
                var firstCellByRow = worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByRows);
                if (firstCellByRow == null)
                    return string.Empty;

                if (foundByFirstColumn)
                {
                    if (firstCellByRow.Column < firstColumn) firstColumn = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstColumn);
                    if (firstCellByRow.Row < firstRow) firstRow = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstRow);
                }
                else
                {
                    firstColumn = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstColumn);
                    firstRow = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstRow);
                }

                Marshal.ReleaseComObject(firstCellByRow);
            }

            foreach (var shapePosition in shapesPosition)
            {
                if (shapePosition.TopLeftColumn < firstColumn)
                    firstColumn = shapePosition.TopLeftColumn;

                if (shapePosition.TopLeftRow < firstRow)
                    firstRow = shapePosition.TopLeftRow;
            }

            var lastColumn = firstColumn;
            var lastRow = firstRow;

            var lastCellByColumn =
                worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByColumns,
                    SearchDirection: ExcelInterop.XlSearchDirection.xlPrevious);

            if (lastCellByColumn != null)
            {
                lastColumn = lastCellByColumn.Column;
                lastRow = lastCellByColumn.Row;
                Marshal.ReleaseComObject(lastCellByColumn);
            }

            var lastCellByRow =
                worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByRows,
                    SearchDirection: ExcelInterop.XlSearchDirection.xlPrevious);

            if (lastCellByRow != null)
            {
                if (lastCellByRow.Column > lastColumn) 
                    lastColumn = CheckForMergedCell(lastCellByRow, MergedCellSearchOrder.LastColumn);

                if (lastCellByRow.Row > lastRow) 
                    lastRow = CheckForMergedCell(lastCellByRow, MergedCellSearchOrder.LastRow);

                var protection = worksheet.Protection;
                if (!worksheet.ProtectContents || protection.AllowDeletingRows)
                {
                    var previousLastCellByRow =
                        worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByRows,
                            SearchDirection: ExcelInterop.XlSearchDirection.xlPrevious,
                            After: lastCellByRow);

                    Marshal.ReleaseComObject(lastCellByRow);

                    if (previousLastCellByRow != null)
                    {
                        var previousRow = CheckForMergedCell(previousLastCellByRow, MergedCellSearchOrder.LastRow);
                        Marshal.ReleaseComObject(previousLastCellByRow);

                        if (previousRow < lastRow - 2)
                        {
                            var rangeToDelete =
                                worksheet.Range[GetColumnAddress(firstColumn) + (previousRow + 1) + ":" +
                                                GetColumnAddress(lastColumn) + (lastRow - 2)];

                            rangeToDelete.Delete(ExcelInterop.XlDeleteShiftDirection.xlShiftUp);
                            Marshal.ReleaseComObject(rangeToDelete);
                            lastRow = previousRow + 2;
                        }
                    }

                    Marshal.ReleaseComObject(protection);
                }
            }

            foreach (var shapePosition in shapesPosition)
            {
                if (shapePosition.BottomRightColumn > lastColumn)
                    lastColumn = shapePosition.BottomRightColumn;

                if (shapePosition.BottomRightRow > lastRow)
                    lastRow = shapePosition.BottomRightRow;
            }

            return GetColumnAddress(firstColumn) + firstRow + ":" +
                   GetColumnAddress(lastColumn) + lastRow;
        }
        #endregion

        #region CountVerticalPageBreaks
        /// <summary>
        /// Returns the total number of vertical pagebreaks in the print area
        /// </summary>
        /// <param name="pageBreaks"></param>
        /// <returns></returns>
        private static int CountVerticalPageBreaks(ExcelInterop.VPageBreaks pageBreaks)
        {
            var result = 0;

            try
            {
                foreach (ExcelInterop.VPageBreak pageBreak in pageBreaks)
                {
                    if (pageBreak.Extent == ExcelInterop.XlPageBreakExtent.xlPageBreakPartial)
                        result += 1;

                    Marshal.ReleaseComObject(pageBreak);
                }
            }
            catch (COMException)
            {
                result = pageBreaks.Count;
            }

            return result;
        }
        #endregion
        
        #region SetWorkSheetPaperSize
        /// <summary>
        /// This method wil figure out the optimal paper size to use and sets it
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="printArea"></param>
        private static void SetWorkSheetPaperSize(ExcelInterop._Worksheet worksheet, string printArea)
        {
            var pageSetup = worksheet.PageSetup;
            var pages = pageSetup.Pages;

            pageSetup.PrintArea = printArea;
            pageSetup.LeftHeader = worksheet.Name;
            
            var pageCount = pages.Count;

            if (pageCount == 1)
                return;

            try
            {
                pageSetup.Order = ExcelInterop.XlOrder.xlOverThenDown;

                foreach (var paperSize in PaperSizes)
                {
                    var exitfor = false;
                    pageSetup.PaperSize = paperSize.PaperSize;
                    pageSetup.Orientation = paperSize.Orientation;
                    worksheet.ResetAllPageBreaks();

                    foreach (var zoomRatio in ZoomRatios)
                    {
                        // Yes these page counts look lame, but so is Excel 2010 in not updating
                        // the pages collection otherwise. We need to call the count methods to
                        // make this code work
                        pageSetup.Zoom = zoomRatio;
                        // ReSharper disable once RedundantAssignment
                        pageCount = pages.Count;

                        if (CountVerticalPageBreaks(worksheet.VPageBreaks) == 0)
                        {
                            exitfor = true;
                            break;
                        }
                    }

                    if (exitfor)
                        break;
                }
            }
            finally
            {
                Marshal.ReleaseComObject(pages);
                Marshal.ReleaseComObject(pageSetup);
            }

        }
        #endregion

        #region SetChartPaperSize
        /// <summary>
        /// This method wil set the papersize for a chart
        /// </summary>
        /// <param name="chart"></param>
        private static void SetChartPaperSize(ExcelInterop._Chart chart)
        {
            var pageSetup = chart.PageSetup;
            var pages = pageSetup.Pages;

            try
            {
                pageSetup.LeftHeader = chart.Name;
                pageSetup.PaperSize = ExcelInterop.XlPaperSize.xlPaperA4;
                pageSetup.Orientation = ExcelInterop.XlPageOrientation.xlLandscape;
            }
            finally
            {
                Marshal.ReleaseComObject(pages);
                Marshal.ReleaseComObject(pageSetup);
            }
        }
        #endregion

        #region Convert
        /// <summary>
        /// Converts an Excel sheet to PDF
        /// </summary>
        /// <param name="inputFile">The Excel input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile"/> has to many rows</exception>
        internal static void Convert(string inputFile, string outputFile)
        {
            // We only need to perform this check if we are running on a server
            if (NativeMethods.IsWindowsServer())
                CheckIfSystemProfileDesktopDirectoryExists();
            CheckIfPrinterIsInstalled();
            DeleteAutoRecoveryFiles();
            string formatString = outputFile.Substring(outputFile.LastIndexOf(".") + 1);

            if (formatString == "txt")
                //ReadAsTXT(inputFile, outputFile);
                SaveAsTXT(inputFile, outputFile);
            if (formatString == "pdf")
                SaveAsPDF(inputFile, outputFile);
            //释放内存
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        //#region GetCon
        ///// <summary>
        ///// Connect to excle
        ///// </summary>
        //private static OleDbConnection GetCon(string excelPath)
        //{
        //    try
        //    {
        //        string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + excelPath + ";" + "Extended Properties=Excel 8.0;";
        //        OleDbConnection conn = new OleDbConnection(strConn);
        //        conn.Open();
        //        return conn;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new ArgumentException("Failed Open Excel", ex.Message);
        //    }
        //}
        //#endregion

        #region GetSheetData
        /// <summary>
        /// Connect to excle
        /// </summary>
        private static DataTable GetSheetData(ExcelInterop.Worksheet sheet)
        {
            DataTable dt = new DataTable();
            if (sheet != null)
            {
                string cellContent;
                int iRowCount = sheet.UsedRange.Rows.Count;
                int iColCount = sheet.UsedRange.Columns.Count;
                ExcelInterop.Range range;
                //负责列头Start
                DataColumn dc;
                int ColumnID = 1;
                range = (ExcelInterop.Range)sheet.Cells[1, 1];
                while (range.Text.ToString().Trim() != "")
                {
                    dc = new DataColumn();
                    dc.DataType = System.Type.GetType("System.String");
                    dc.ColumnName = range.Text.ToString().Trim();
                    dt.Columns.Add(dc);
                    range = (ExcelInterop.Range)sheet.Cells[1, ++ColumnID];
                }
                for (int iRow = 2; iRow <= iRowCount; iRow++)
                {
                    DataRow dr = dt.NewRow();
                    for (int iCol = 1; iCol <= iColCount; iCol++)
                    {
                        range = (ExcelInterop.Range)sheet.Cells[iRow, iCol];
                        cellContent = (range.Value2 == null) ? "" : range.Text.ToString();
                        dr[iCol-1] = cellContent;
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }
        #endregion

        #region GetSheetData
        /// <summary>
        /// Connect to excle
        /// </summary>
        private static void SaveDataSet(DataSet ds,string outputFile)
        {
            StreamWriter mywrite = new StreamWriter(outputFile, false, System.Text.Encoding.UTF8, 100);
            for (int t = 0; t < ds.Tables.Count; t++)
            {
                for (int i = 0; i < ds.Tables[t].Columns.Count; i++)
                {
                    mywrite.Write(ds.Tables[t].Columns[i].Caption + "\t");
                }
                mywrite.WriteLine();
                for (int j = 0; j < ds.Tables[t].Rows.Count; j++)
                {
                    for (int k = 0; k < ds.Tables[t].Columns.Count; k++)
                    {
                        mywrite.Write(ds.Tables[t].Rows[j].ItemArray.GetValue(k) + "\t");
                    }
                    mywrite.WriteLine();
                }
                mywrite.WriteLine();
            }
            mywrite.Close();
        }
        #endregion

        #region ReadAsTXT
        /// <summary>
        /// 将Excel文件另存为TXT格式
        /// </summary>
        private static void ReadAsTXT(string inputFile, string outputFile)
        {
            ExcelInterop.Application excel = null;
            ExcelInterop.Workbook workbook = null;
            string tempFileName = null;
            DataSet ds = new DataSet();
            try
            {
                excel = new ExcelInterop.ApplicationClass
                {
                    ScreenUpdating = false,
                    DisplayAlerts = false,
                    DisplayDocumentInformationPanel = false,
                    DisplayRecentFiles = false,
                    DisplayScrollBars = false,
                    AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable,
                    PrintCommunication = true, // DO NOT REMOVE THIS LINE, NO NEVER EVER ... DON'T EVEN TRY IT
                    Visible = false
                };

                var extension = Path.GetExtension(inputFile);

                if (string.IsNullOrWhiteSpace(extension))
                    extension = string.Empty;

                workbook = Open(excel, inputFile, extension, false);

                // We cannot determine a print area when the document is marked as final so we remove this
                //workbook.Final = false;

                // Fix for "This command is not available in a shared workbook."
                if (workbook.MultiUserEditing)
                {
                    tempFileName = Path.GetTempFileName() + Guid.NewGuid() + Path.GetExtension(inputFile);
                    workbook.SaveAs(tempFileName, AccessMode: ExcelInterop.XlSaveAsAccessMode.xlExclusive);
                }
                foreach (var sheetObject in workbook.Sheets)
                {
                    var sheet = sheetObject as ExcelInterop.Worksheet;
                    DataTable dt = new DataTable();
                    dt = GetSheetData(sheet);
                    ds.Tables.Add(dt);
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }
                //读取ds数据集，保存到txt
                SaveDataSet(ds, outputFile);
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Saved = true;
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                }

                if (!string.IsNullOrEmpty(tempFileName) && File.Exists(tempFileName))
                    File.Delete(tempFileName);
            }
        }
        #endregion


        #region SaveAsTXT
        /// <summary>
        /// 将Excel文件另存为PDF格式
        /// </summary>
        private static void SaveAsTXT(string inputFile, string outputFile)
        {
            ExcelInterop.Application excel = null;
            ExcelInterop.Workbook workbook = null;
            string tempFileName = "";
            try
            {
                excel = new ExcelInterop.ApplicationClass
                {
                    ScreenUpdating = false,
                    DisplayAlerts = false,
                    DisplayDocumentInformationPanel = false,
                    DisplayRecentFiles = false,
                    DisplayScrollBars = false,
                    AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable,
                    PrintCommunication = true, // DO NOT REMOVE THIS LINE, NO NEVER EVER ... DON'T EVEN TRY IT
                    Visible = false
                };

                var extension = Path.GetExtension(inputFile);

                if (string.IsNullOrWhiteSpace(extension))
                    extension = string.Empty;

                workbook = Open(excel, inputFile, extension, false);
                
                workbook.Final = false;
                
                if (workbook.MultiUserEditing)
                {
                    tempFileName = Path.GetTempFileName() + Guid.NewGuid() + Path.GetExtension(inputFile);
                    workbook.SaveAs(tempFileName, AccessMode: ExcelInterop.XlSaveAsAccessMode.xlExclusive);
                }

                var usedSheets = 0;

                foreach (var sheetObject in workbook.Sheets)
                {
                    var sheet = sheetObject as ExcelInterop.Worksheet;

                    if (sheet != null)
                    {
                        var protection = sheet.Protection;
                        var activeWindow = excel.ActiveWindow;

                        try
                        {
                            // ReSharper disable once RedundantCast
                            ((Microsoft.Office.Interop.Excel._Worksheet)sheet).Activate();
                            if (!sheet.ProtectContents || protection.AllowFormattingColumns)
                                if (activeWindow.View != ExcelInterop.XlWindowView.xlPageLayoutView)
                                    sheet.Columns.AutoFit();

                        }
                        catch (COMException)
                        {
                            // Do nothing, this sometimes failes and there is nothing we can do about it
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(activeWindow);
                            Marshal.ReleaseComObject(protection);
                        }

                        var printArea = GetWorksheetPrintArea(sheet);

                        switch (printArea)
                        {
                            case "shapes":
                                SetWorkSheetPaperSize(sheet, string.Empty);
                                usedSheets += 1;
                                break;

                            case "":
                                break;

                            default:
                                SetWorkSheetPaperSize(sheet, printArea);
                                usedSheets += 1;
                                break;
                        }

                        Marshal.ReleaseComObject(sheet);
                        continue;
                    }

                    var chart = sheetObject as ExcelInterop.Chart;
                    if (chart != null)
                    {
                        SetChartPaperSize(chart);
                        Marshal.ReleaseComObject(chart);
                    }
                }

                // It is not possible in Excel to export an empty workbook
                if (usedSheets != 0)
                {
                    workbook.SaveAs(outputFile, ExcelInterop.XlFileFormat.xlUnicodeText);
                    //FileStream fs = new FileStream(outputFile, FileMode.Open, FileAccess.ReadWrite);
                    //StreamReader sr = new StreamReader(fs);
                    //string str = sr.ReadToEnd();
                    ////while (str != null)
                    ////{
                    ////    str += sr.ReadLine()+"\n";
                    ////}
                    //sr.Close();
                    //StreamWriter sw = new StreamWriter(outputFile, false, System.Text.Encoding.UTF8, 100);
                    //sw.Write(str);
                    //sw.Close();
                    //fs.Close();
                }
                else
                {
                    //修改空字段
                    FileStream fs = new FileStream(outputFile, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    StreamWriter sw = new StreamWriter(fs);
                    sw.Write("Excle is Null");
                    sw.Close();
                    fs.Close();
                }
                    
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Saved = true;
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }

                if (excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                }

                if (!string.IsNullOrEmpty(tempFileName) && File.Exists(tempFileName))
                    File.Delete(tempFileName);

                FileStream fs = new FileStream(outputFile, FileMode.Open, FileAccess.ReadWrite);
                StreamReader sr = new StreamReader(fs);
                string str = sr.ReadToEnd();
                sr.Close();
                StreamWriter sw = new StreamWriter(outputFile, false, System.Text.Encoding.UTF8, 100);
                sw.Write(str);
                sw.Close();
                fs.Close();
            }
        }
        #endregion






        #region SaveAsPDF
        /// <summary>
        /// 将Excel文件另存为PDF格式
        /// </summary>
        private static void SaveAsPDF(string inputFile, string outputFile)
        {
            ExcelInterop.Application excel = null;
            ExcelInterop.Workbook workbook = null;
            string tempFileName = null;
            try
            {
                excel = new ExcelInterop.ApplicationClass
                {
                    ScreenUpdating = false,
                    DisplayAlerts = false,
                    DisplayDocumentInformationPanel = false,
                    DisplayRecentFiles = false,
                    DisplayScrollBars = false,
                    AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable,
                    PrintCommunication = true, // DO NOT REMOVE THIS LINE, NO NEVER EVER ... DON'T EVEN TRY IT
                    Visible = false
                };

                var extension = Path.GetExtension(inputFile);

                if (string.IsNullOrWhiteSpace(extension))
                    extension = string.Empty;

                if (extension.ToUpperInvariant() == ".CSV")
                {
                    // Yes this look somewhat weird but we have to change the extension if we want to handle
                    // CSV files with different kind of separators. Otherwhise Excel will always overrule whatever
                    // setting we make to open a file
                    tempFileName = Path.GetTempFileName() + Guid.NewGuid() + ".txt";
                    File.Copy(inputFile, tempFileName);
                    inputFile = tempFileName;
                }

                workbook = Open(excel, inputFile, extension, false);

                // We cannot determine a print area when the document is marked as final so we remove this
                workbook.Final = false;

                // Fix for "This command is not available in a shared workbook."
                if (workbook.MultiUserEditing)
                {
                    tempFileName = Path.GetTempFileName() + Guid.NewGuid() + Path.GetExtension(inputFile);
                    workbook.SaveAs(tempFileName, AccessMode: ExcelInterop.XlSaveAsAccessMode.xlExclusive);
                }

                var usedSheets = 0;

                foreach (var sheetObject in workbook.Sheets)
                {
                    var sheet = sheetObject as ExcelInterop.Worksheet;

                    if (sheet != null)
                    {
                        var protection = sheet.Protection;
                        var activeWindow = excel.ActiveWindow;

                        try
                        {
                            // ReSharper disable once RedundantCast
                            ((Microsoft.Office.Interop.Excel._Worksheet)sheet).Activate();
                            if (!sheet.ProtectContents || protection.AllowFormattingColumns)
                                if (activeWindow.View != ExcelInterop.XlWindowView.xlPageLayoutView)
                                    sheet.Columns.AutoFit();

                        }
                        catch (COMException)
                        {
                            // Do nothing, this sometimes failes and there is nothing we can do about it
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(activeWindow);
                            Marshal.ReleaseComObject(protection);
                        }

                        var printArea = GetWorksheetPrintArea(sheet);

                        switch (printArea)
                        {
                            case "shapes":
                                SetWorkSheetPaperSize(sheet, string.Empty);
                                usedSheets += 1;
                                break;

                            case "":
                                break;

                            default:
                                SetWorkSheetPaperSize(sheet, printArea);
                                usedSheets += 1;
                                break;
                        }

                        Marshal.ReleaseComObject(sheet);
                        continue;
                    }

                    var chart = sheetObject as ExcelInterop.Chart;
                    if (chart != null)
                    {
                        SetChartPaperSize(chart);
                        Marshal.ReleaseComObject(chart);
                    }
                }

                // It is not possible in Excel to export an empty workbook
                if (usedSheets != 0)
                {
                    workbook.ExportAsFixedFormat(ExcelInterop.XlFixedFormatType.xlTypePDF, outputFile);
                }
                else
                    throw new OCFileContainsNoData("The file '" + Path.GetFileName(inputFile) + "' contains no data");
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Saved = true;
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }

                if (excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                }

                if (!string.IsNullOrEmpty(tempFileName) && File.Exists(tempFileName))
                    File.Delete(tempFileName);
            }
        }
        #endregion





        #region IsPasswordProtected
        /// <summary>
        /// Returns true when the Excel file is password protected
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// <exception cref="OCFileIsCorrupt">Raised when the file is corrupt</exception>
        public static bool IsPasswordProtected(string fileName)
        {
            try
            {
                using (var compoundFile = new CompoundFile(fileName))
                {
                    if (compoundFile.RootStorage.TryGetStream("EncryptedPackage") != null) return true;

                    var stream = compoundFile.RootStorage.TryGetStream("WorkBook");
                    if (stream == null)
                        compoundFile.RootStorage.TryGetStream("Book");

                    if (stream == null)
                        throw new OCFileIsCorrupt("Could not find the WorkBook or Book stream in the file '" + fileName +
                                                  "'");

                    var bytes = stream.GetData();
                    using (var memoryStream = new MemoryStream(bytes))
                    using (var binaryReader = new BinaryReader(memoryStream))
                    {
                        // Get the record type, at the beginning of the stream this should always be the BOF
                        var recordType = binaryReader.ReadUInt16();

                        // Something seems to be wrong, we would expect a BOF but for some reason it isn't so stop it
                        if (recordType != 0x809)
                            throw new OCFileIsCorrupt("The file '" + fileName + "' is corrupt");

                        var recordLength = binaryReader.ReadUInt16();
                        binaryReader.BaseStream.Position += recordLength;

                        // Search after the BOF for the FilePass record, this starts with 2F hex
                        recordType = binaryReader.ReadUInt16();
                        return recordType == 0x2F;
                    }
                }
            }
            catch (CFCorruptedFileException)
            {
                throw new OCFileIsCorrupt("The file '" + Path.GetFileName(fileName) + "' is corrupt");
            }
            catch (CFFileFormatException)
            {
                // It seems the file is just a normal Microsoft Office 2007 and up Open XML file
                return false;
            }
        }
        #endregion

        #region GetCsvSeperator
        /// <summary>
        /// Returns the seperator and textqualifier that is used in the CSV file
        /// </summary>
        /// <param name="inputFile">The inputfile</param>
        /// <param name="separator">The separator that is used</param>
        /// <param name="textQualifier">The text qualifier</param>
        /// <returns></returns>
        private static void GetCsvSeperator(string inputFile, out string separator, out ExcelInterop.XlTextQualifier textQualifier)
        {
            separator = string.Empty;
            textQualifier = ExcelInterop.XlTextQualifier.xlTextQualifierNone;

            using (var streamReader = new StreamReader(inputFile))
            {
                var line = string.Empty;
                while (string.IsNullOrEmpty(line))
                    line = streamReader.ReadLine();

                if (line.Contains(";")) separator = ";";
                else if (line.Contains(",")) separator = ",";
                else if (line.Contains("\t")) separator = "\t";
                else if (line.Contains(" ")) separator = " ";

                if (line.Contains("\"")) textQualifier = ExcelInterop.XlTextQualifier.xlTextQualifierDoubleQuote;
                else if (line.Contains("'")) textQualifier = ExcelInterop.XlTextQualifier.xlTextQualifierSingleQuote;
            }
        }
        #endregion

        #region Open
        /// <summary>
        /// Opens the <paramref name="inputFile"/> and returns it as an <see cref="ExcelInterop.Workbook"/> object
        /// </summary>
        /// <param name="excel">The <see cref="ExcelInterop.Application"/></param>
        /// <param name="inputFile">The file to open</param>
        /// <param name="extension">The file extension</param>
        /// <param name="repairMode">When true the <paramref name="inputFile"/> is opened in repair mode</param>
        /// <returns></returns>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile"/> has to many rows</exception>
        private static ExcelInterop.Workbook Open(ExcelInterop._Application excel,
                                                   string inputFile,
                                                   string extension,
                                                   bool repairMode)
        {
            try
            {
                switch (extension.ToUpperInvariant())
                {
                    case ".CSV":

                        var count = File.ReadLines(inputFile).Count();
                        var excelMaxRows = MaxRows;
                        if (count > excelMaxRows)
                            throw new OCCsvFileLimitExceeded("The input CSV file has more then " + excelMaxRows +
                                                             " rows, the installed Excel version supports only " +
                                                             excelMaxRows + " rows");

                        string separator;
                        ExcelInterop.XlTextQualifier textQualifier;

                        GetCsvSeperator(inputFile, out separator, out textQualifier);

                        switch (separator)
                        {
                            case ";":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited,
                                    textQualifier, true, false, true);
                                return excel.ActiveWorkbook;

                            case ",":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, false, true);
                                return excel.ActiveWorkbook;

                            case "\t":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, true);
                                return excel.ActiveWorkbook;

                            case " ":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, false, false, true);
                                return excel.ActiveWorkbook;

                            default:
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, true);
                                return excel.ActiveWorkbook;
                        }

                    default:

                        if (repairMode)
                            return excel.Workbooks.Open(inputFile, false, true,
                                Password: "dummypassword",
                                IgnoreReadOnlyRecommended: true,
                                AddToMru: false,
                                CorruptLoad: ExcelInterop.XlCorruptLoad.xlRepairFile);

                        return excel.Workbooks.Open(inputFile, false, true,
                            Password: "dummypassword",
                            IgnoreReadOnlyRecommended: true,
                            AddToMru: false);

                }
            }
            catch (COMException comException)
            {
                if (comException.ErrorCode == -2146827284)
                    throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                        "' is password protected");

                throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                                        "' could not be opened, error: " + ExceptionHelpers.GetInnerException(comException));
            }
            catch (Exception exception)
            {
                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                              "' could not be opened, error: " +
                                              ExceptionHelpers.GetInnerException(exception));

                return Open(excel, inputFile, extension, true);
            }
        }
        #endregion

        #region DeleteAutoRecoveryFiles
        /// <summary>
        /// This method will delete the automatic created Resiliency key. Excel uses this registry key  
        /// to make entries to corrupted workbooks. If there are to many entries under this key Excel will
        /// get slower and slower to start. To prevent this we just delete this key when it exists
        /// </summary>
        private static void DeleteAutoRecoveryFiles()
        {
            try
            {
                // HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Excel\Resiliency\DocumentRecovery
                var version = string.Empty;

                switch (VersionNumber)
                {
                    // Word 2003
                    case 11:
                        version = "11.0";
                        break;

                    // Word 2017
                    case 12:
                        version = "12.0";
                        break;

                    // Word 2010
                    case 14:
                        version = "14.0";
                        break;

                    // Word 2013
                    case 15:
                        version = "15.0";
                        break;

                    // Word 2016
                    case 16:
                        version = "16.0";
                        break;
                }

                var key = @"Software\Microsoft\Office\" + version + @"\Excel\Resiliency";

                if (Registry.CurrentUser.OpenSubKey(key, false) != null)
                    Registry.CurrentUser.DeleteSubKeyTree(key);
            }
            catch (Exception exception)
            {
                EventLog.WriteEntry("OfficeConverter", ExceptionHelpers.GetInnerException(exception), EventLogEntryType.Error);
            }
        }
        #endregion
    }
}
