using System;
using System.IO;
using System.Linq;
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
            const string FILE_EXTENSION = ".xlsx";
            // Create Spreadsheet encapsulation
            SLDocument oSLDocument = new SLDocument();
            System.Data.DataTable dt = new System.Data.DataTable();
            Console.WriteLine("Creating data table...");
            // Add columns to DataTable
            dt = AddDataTableCols(dt);
            // Add rows to DataTable
            dt = AddDataTableRows(filesArray, dt);
            // Import DataTable to spreadsheet
            Console.WriteLine("Importing data table to spreadsheet...");
            oSLDocument.ImportDataTable(1, 1, dt, true);
            Console.Write("Insert the name of the spreadsheet file: ");
            string userInputFileName = Console.ReadLine();
            string newFilePath = path + "\\" + userInputFileName + FILE_EXTENSION;

            // Checking If the name introduced by the user contains invalid characters
            while(CheckFileNameInvalidChars(userInputFileName))
            {
                Console.WriteLine("Please, do not select a name with invalid characters for the spreadsheet file");
                Console.Write("Insert a new name for the spreadsheet file: ");
                userInputFileName = Console.ReadLine();
                newFilePath = path + "\\" + userInputFileName + FILE_EXTENSION;
            }

            // Checking If file exists
            if (CheckIfFileExists(newFilePath))
            {
                // Overwrite the file
                Console.WriteLine("There's already a spreadsheet file with the given name and the extension " + FILE_EXTENSION);
                Console.Write("Do you want to overwrite the file? (Y/N): ");
                string overwriteFileName = Convert.ToString(Console.ReadLine());
                if (overwriteFileName.ToUpper() == "N" || overwriteFileName.ToUpper() == "NO")
                {
                    newFilePath = CreateNameCopy(path, userInputFileName, FILE_EXTENSION);
                }
            }
            Console.Write("Saving file...");
            oSLDocument.SaveAs(newFilePath);
            Console.WriteLine("Done");
        }

        public static System.Data.DataTable AddDataTableCols(System.Data.DataTable pDataTable)
        {
            System.Data.DataTable returnDataTable = pDataTable;
            returnDataTable.Columns.Add("File Name", typeof(string));
            bool addNewCol = true;
            while (addNewCol)
            {
                Console.Write("Do you want to add another column? (Y/N): ");
                string moreColumns = Convert.ToString(Console.ReadLine());
                if (moreColumns.ToUpper() == "Y")
                {
                    Console.Write("Type the name of the column: ");
                    string newCol = Console.ReadLine();
                    returnDataTable.Columns.Add(newCol, typeof(string));
                    Console.SetCursorPosition(0, Console.CursorTop - 1);
                    Console.Write(new string(' ', Console.WindowWidth));
                }
                else
                    addNewCol = false;
                Console.SetCursorPosition(0, Console.CursorTop - 1);
                Console.Write(new string(' ', Console.WindowWidth));
                Console.SetCursorPosition(0, Console.CursorTop);
            };
            return returnDataTable;
        }

        public static System.Data.DataTable AddDataTableRows(string[] pFilesArray, System.Data.DataTable pDataTable)
        {
            System.Data.DataTable returnDataTable = pDataTable;
            int filesArrayLength = pFilesArray.Length;
            int filesCount = 1;
            int currentLineCursor = 0;

            // Hide cursor
            Console.CursorVisible = false;
            // Loading files
            string loadingText = "Loading ";
            Console.Write(loadingText);
            // Writting the loading message
            Console.Write("{0}", filesCount.ToString("D" + filesArrayLength.ToString().Length));
            Console.Write(string.Format("/{0}", filesArrayLength));
            foreach (string file in pFilesArray)
            {
                if (filesCount > 1)
                {
                    currentLineCursor = Console.CursorTop;
                    Console.SetCursorPosition(loadingText.Length, currentLineCursor);
                    Console.Write("{0}", filesCount.ToString("D" + filesArrayLength.ToString().Length));
                }
                filesCount++;
                // Add file to rows
                pDataTable.Rows.Add(Path.GetFileName(file));
            }
            // Show Cursor
            Console.CursorVisible = true;
            Console.WriteLine();
            return pDataTable;
        }

        public static string CreateNameCopy (string pDirectoryPath, string pFileName, string pFILE_EXTENSION)
        {
            string copyWatermark = "_copy_";
            int copyNumber = 0;
            string newFileName, newPath;
            bool copyExists;
            do
            {
                copyNumber++;
                newFileName = pFileName + copyWatermark + copyNumber;
                newPath = pDirectoryPath + "\\" + newFileName + pFILE_EXTENSION;
                copyExists = CheckIfFileExists(newPath);
            } while (copyExists);
            return newPath;
        }
        
        public static bool CheckFileNameInvalidChars(string pFileName)
        {
            return pFileName.Any(Path.GetInvalidFileNameChars().Contains);
        }

        public static bool CheckIfFileExists(string filePath)
        {
            return File.Exists(filePath);
        }
    }
}
