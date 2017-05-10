﻿using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Win32;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;
using OpenMcdf;
using WordInterop = Microsoft.Office.Interop.Word;

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
    /// This class is used as a placeholder for al Word related methods
    /// </summary>
    internal static class Word
    {
        #region Fields
        /// <summary>
        /// Word version number
        /// </summary>
        private static readonly int VersionNumber;
        #endregion

        #region Constructor
        /// <summary>
        /// This constructor is called the first time when the <see cref="Convert"/> or
        /// <see cref="IsPasswordProtected"/> method is called. Some checks are done to
        /// see if all requirements for a succesfull conversion are there.
        /// </summary>
        /// <exception cref="OCConfiguration">Raised when the registry could not be read to determine Word version</exception>
        static Word()
        {
            try
            {
                var baseKey = Registry.ClassesRoot;
                var subKey = baseKey.OpenSubKey(@"Word.Application\CurVer");
                if (subKey != null)
                {
                    switch (subKey.GetValue(string.Empty).ToString().ToUpperInvariant())
                    {
                        // Word 2003
                        case "WORD.APPLICATION.11":
                            VersionNumber = 11;
                            break;

                        // Word 2007
                        case "WORD.APPLICATION.12":
                            VersionNumber = 12;
                            break;

                        // Word 2010
                        case "WORD.APPLICATION.14":
                            VersionNumber = 14;
                            break;

                        // Word 2013
                        case "WORD.APPLICATION.15":
                            VersionNumber = 15;
                            break;

                        // Word 2016
                        case "WORD.APPLICATION.16":
                            VersionNumber = 16;
                            break;

                        default:
                            throw new OCConfiguration("Could not determine WORD version");
                    }
                }
                else
                    throw new OCConfiguration("Could not find registry key WORD.Application\\CurVer");
            }
            catch (Exception exception)
            {
                throw new OCConfiguration("Could not read registry to check WORD version", exception);
            }
        }
        #endregion

        #region Convert
        /// <summary>
        /// Converts a Word document to PDF
        /// </summary>
        /// <param name="inputFile">The Word input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        internal static void Convert(string inputFile, string outputFile)
        {
            DeleteAutoRecoveryFiles();

            WordInterop.ApplicationClass word = null;
            WordInterop.DocumentClass document = null;

            try
            {
                //select output format
                object format = WordInterop.WdSaveFormat.wdFormatDocumentDefault;
                string formatString = outputFile.Substring(outputFile.LastIndexOf(".") + 1);
                if (formatString == "doc")
                    format = WordInterop.WdSaveFormat.wdFormatDocument;
                if (formatString == "docx")
                    format = WordInterop.WdSaveFormat.wdFormatXMLDocument;
                if (formatString == "rtf")
                    format = WordInterop.WdSaveFormat.wdFormatRTF;
                if (formatString == "txt")
                    format = WordInterop.WdSaveFormat.wdFormatText;
                if (formatString == "html")
                    format = WordInterop.WdSaveFormat.wdFormatHTML;
                if (formatString == "odt")
                    format = WordInterop.WdSaveFormat.wdFormatOpenDocumentText;
                if (formatString == "pdf")
                    format = WordInterop.WdSaveFormat.wdFormatPDF;
                if (formatString == "xps")
                    format = WordInterop.WdSaveFormat.wdFormatXPS;
                if (formatString == "xml")
                    format = WordInterop.WdSaveFormat.wdFormatXML;
                if (formatString == "unicode")
                    format = WordInterop.WdSaveFormat.wdFormatUnicodeText;
                
                word = new WordInterop.ApplicationClass
                {
                    ScreenUpdating = false,
                    DisplayAlerts = WordInterop.WdAlertLevel.wdAlertsNone,
                    DisplayDocumentInformationPanel = false,
                    DisplayRecentFiles = false,
                    DisplayScrollBars = false,
                    AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
                };

                word.Options.UpdateLinksAtOpen = false;
                word.Options.ConfirmConversions = false;
                word.Options.SaveInterval = 0;
                word.Options.SaveNormalPrompt = false;
                word.Options.SavePropertiesPrompt = false;
                word.Options.AllowReadingMode = false;
                word.Options.WarnBeforeSavingPrintingSendingMarkup = false;
                word.Options.UpdateFieldsAtPrint = false;
                word.Options.UpdateLinksAtOpen = false;
                word.Options.UpdateLinksAtPrint = false;
                
                document = (WordInterop.DocumentClass) Open(word, inputFile, false);

                // Do not remove this line!!
                // This is yet another solution to a weird Office problem. Sometimes there
                // are Word documents with images in it that take some time to load. When
                // we remove the line below the ExportAsFixedFormat method will be called 
                // before the images are loaded thus resulting in an unendless loop somewhere
                // in this method.
                // ReSharper disable once UnusedVariable
                var count = document.ComputeStatistics(WordInterop.WdStatistic.wdStatisticPages);

                word.DisplayAutoCompleteTips = false;
                word.DisplayScreenTips = false;
                word.DisplayStatusBar = false;
                document.SaveAs2(outputFile, format,Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, MsoEncoding.msoEncodingUTF8);

                //document.SaveAs2(outputFile,format);
                //保存后转码
            }
            finally
            {
                if (document != null)
                {
                    document.Saved = true;
                    document.Close(false);
                    Marshal.ReleaseComObject(document);
                }

                if (word != null)
                {
                    word.Quit(false);
                    Marshal.ReleaseComObject(word);
                }
            }
        }
        #endregion

        #region IsPasswordProtected
        /// <summary>
        /// Returns true when the Word file is password protected
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

                    var stream = compoundFile.RootStorage.TryGetStream("WordDocument");

                    if (stream == null)
                        throw new OCFileIsCorrupt("Could not find the WordDocument stream in the file '" + fileName +
                                                  "'");

                    var bytes = stream.GetData();
                    using (var memoryStream = new MemoryStream(bytes))
                    using (var binaryReader = new BinaryReader(memoryStream))
                    {
                        //http://msdn.microsoft.com/en-us/library/dd944620%28v=office.12%29.aspx
                        // The bit that shows if the file is encrypted is in the 11th and 12th byte so we 
                        // need to skip the first 10 bytes
                        binaryReader.ReadBytes(10);

                        // Now we read the 2 bytes that we need
                        var pnNext = binaryReader.ReadUInt16();
                        //(value & mask) == mask)

                        // The bit that tells us if the file is encrypted
                        return (pnNext & 0x0100) == 0x0100;
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

        #region Open
            /// <summary>
            /// Opens the <paramref name="inputFile"/> and returns it as an <see cref="WordInterop.Document"/> object
            /// </summary>
            /// <param name="word">The <see cref="WordInterop.Application"/></param>
            /// <param name="inputFile">The file to open</param>
            /// <param name="repairMode">When true the <paramref name="inputFile"/> is opened in repair mode</param>
            /// <returns></returns>
        private static WordInterop.Document Open(WordInterop._Application word,
                                                string inputFile,
                                                bool repairMode)
        {
            try
            {
                WordInterop.Document document;

                var extension = Path.GetExtension(inputFile);

                if (extension != null && extension.ToUpperInvariant() == ".TXT")
                    document = word.Documents.OpenNoRepairDialog(inputFile, false, true, false, "dummypassword",
                        Format: WordInterop.WdOpenFormat.wdOpenFormatUnicodeText,
                        OpenAndRepair: repairMode,
                        NoEncodingDialog: true);
                else
                    document = word.Documents.OpenNoRepairDialog(inputFile, false, true, false, "dummypassword",
                        OpenAndRepair: repairMode,
                        NoEncodingDialog: true);

                // This will lock or unlock all form fields in a Word document so that auto fill 
                // and date/time field do or don't get updated automaticly when converting
                if (document.Fields.Count > 0)
                {
                    foreach (WordInterop.Field field in document.Fields)
                        field.Locked = true;
                }

                return document;
            }
            catch (Exception exception)
            {
                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                              "' seems to be corrupt, error: " +
                                              ExceptionHelpers.GetInnerException(exception));

                return Open(word, inputFile, true);
            }
        }
        #endregion

        #region DeleteAutoRecoveryFiles
        /// <summary>
        /// This method will delete the automatic created Resiliency key. Word uses this registry key  
        /// to make entries to corrupted documents. If there are to many entries under this key Word will
        /// get slower and slower to start. To prevent this we just delete this key when it exists
        /// </summary>
        private static void DeleteAutoRecoveryFiles()
        {
            try
            {
                // HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Resiliency\DocumentRecovery
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

                var key = @"Software\Microsoft\Office\" + version + @"\Word\Resiliency";

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
