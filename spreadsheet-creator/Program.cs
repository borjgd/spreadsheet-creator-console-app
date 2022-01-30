using System;
using System.IO;
using SpreadsheetLight;

namespace spreadsheet_creator
{
    class Program
    {
        static void Main(string[] args)
        {
            string tmpPath = Environment.GetFolderPath(Environment.SpecialFolder.MyMusic);
            GetAllFiles(tmpPath);
            Console.WriteLine();
            Console.WriteLine("Press any key to finish");
            Console.ReadLine();
        }

        public static void GetAllFiles(string path)
        {
            if (Directory.Exists(path))
            {
                string[] pathExistingFiles = Directory.GetFiles(path);
                if (pathExistingFiles.Length > 0)
                {
                    CreateXLSXFile(path, pathExistingFiles);
                }
                else
                {
                    Console.WriteLine("No files found in this directory");
                }
            }
        }

        public static void CreateXLSXFile(string path, string[] filesArray)
        {
            SLDocument oSLDocument = new SLDocument();
            System.Data.DataTable dt = new System.Data.DataTable();
            Console.WriteLine("Creating data table...");
            // Columns - Headers
            dt.Columns.Add("File Name", typeof(string));
            // Rows - Data
            foreach (string test in filesArray)
            {
                dt.Rows.Add(Path.GetFileName(test));
            }
            Console.WriteLine("Importing data table to spreadsheet...");
            oSLDocument.ImportDataTable(1, 1, dt, true);
            Console.WriteLine("Insert the name of the spreadsheet file: ");
            string userInput_fileName = Console.ReadLine();
            Console.WriteLine("Saving file...");
            oSLDocument.SaveAs(path + "\\" + userInput_fileName + ".xlsx");
            Console.WriteLine("Done");
        }
    }
}
