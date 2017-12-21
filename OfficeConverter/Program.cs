using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace OfficeConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            //args[0] = "D:\\OfficeConverter\\File\\PSW.xlsx";
            //args[1] = "D:\\OfficeConverter\\File\\PSW.txt";
            if (args.Length < 2 || args.Length > 3)
            {
                System.Console.WriteLine("Error: no filenames specified");
                System.Console.WriteLine("Usage: mswordconverter inputfile outputfile");
                System.Console.WriteLine("saveformat is optional, will try to autodetect format and fall back to word default if not given.");
                System.Console.WriteLine("can be one of: doc, docx, rtf, txt, html, odt, pdf, xps, xml, unicode");
                System.Console.WriteLine("format support depends on whether support is present in office itself");
                return;
            }
            string inputFile = args[0];
            string outputFile = args[1];

            Converter converter = new Converter();
            try
            {
                converter.Convert(inputFile, outputFile);
            }
            catch
            {
                System.Environment.Exit(0);
            }
        }
    }
}
