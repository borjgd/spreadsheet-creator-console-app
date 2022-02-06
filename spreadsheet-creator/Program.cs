using System;
using System.IO;
using SpreadsheetLight;

namespace spreadsheet_creator
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                GetAllFiles(args[0]);
            }
            else
            {
                Console.WriteLine("No path passed as argument to the application");
            }
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
            int filesArrayLength = filesArray.Length;
            int count = 1;
            int currentLineCursor = 0;
            SLDocument oSLDocument = new SLDocument();
            System.Data.DataTable dt = new System.Data.DataTable();
            Console.WriteLine("Creating data table...");
            // Columns - Headers
            dt.Columns.Add("File Name", typeof(string));
            
            bool addNewCol = true;
            while(addNewCol)
            {
                Console.Write("Do you want to add another column? (Y/N): ");
                string moreColumns = Convert.ToString(Console.ReadLine());
                if (moreColumns.ToUpper() == "Y")
                {
                    Console.Write("Type the name of the column: ");
                    string newCol = Console.ReadLine();
                    dt.Columns.Add(newCol, typeof(string));
                    Console.SetCursorPosition(0, Console.CursorTop -1);
                    Console.Write(new string(' ', Console.WindowWidth));
                }
                else
                    addNewCol = false;
                Console.SetCursorPosition(0, Console.CursorTop - 1);
                Console.Write(new string(' ', Console.WindowWidth));
                Console.SetCursorPosition(0, Console.CursorTop);
            };
            // Hide cursor
            Console.CursorVisible = false;
            // Rows - Data
            // Loading message
            string loadingText = "Loading ";
            Console.Write(loadingText);
            Console.Write("{0}", count.ToString("D" + filesArrayLength.ToString().Length));
            Console.Write(string.Format("/{0}", filesArrayLength));
            foreach (string test in filesArray)
            {
                if (count > 1)
                {
                    currentLineCursor = Console.CursorTop;
                    Console.SetCursorPosition(loadingText.Length, currentLineCursor);
                    Console.Write("{0}", count.ToString("D" + filesArrayLength.ToString().Length));
                }
                count++;
                // Add file to rows
                dt.Rows.Add(Path.GetFileName(test));
            }
            // Show Cursor
            Console.CursorVisible = true;
            Console.WriteLine();
            // Import datatable to spreadsheet
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
