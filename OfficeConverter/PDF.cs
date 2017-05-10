using System;
using System.IO;
using System.Text;
using org.pdfbox.pdmodel;
using org.pdfbox.util;

namespace OfficeConverter
{
    class PDF
    {
        public static void Converter(FileInfo pdffile, FileInfo txtfile)
        {

            PDDocument doc = PDDocument.load(pdffile.FullName);

            PDFTextStripper pdfStripper = new PDFTextStripper();

            string text = pdfStripper.getText(doc);

            StreamWriter swPdfChange = new StreamWriter(txtfile.FullName, false, Encoding.GetEncoding("utf-8"));

            swPdfChange.Write(text);

            swPdfChange.Close();

        }
    }
}
