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
                var folderPath = args[0];
                var folderFiles = GetAllFiles(folderPath);
                if (folderFiles.Length > 0)
                    CreateSpreadsheet(folderPath, folderFiles);
                else
                    Console.WriteLine("No files found in this directory");
            }
            else
            {
                Console.WriteLine("No path passed as argument to the application");
            }
            Console.WriteLine();
            Console.WriteLine("Press any key to finish");
            Console.ReadLine();
        }

        /// <summary>
        /// Gets all files from the given folder path
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns></returns>
        public static string[] GetAllFiles(string folderPath)
        {
            return Directory.GetFiles(folderPath);
        }

        /// <summary>
        /// Creates spreadsheet adding columns and rows
        /// </summary>
        /// <param name="folderPath"></param>
        /// <param name="filesArray"></param>
        public static void CreateSpreadsheet(string folderPath, string[] filesArray)
        {
            Console.WriteLine("Creating spreadsheet...");
            Spreadsheet spreadsheet = new();
            System.Data.DataTable dt = new();
            dt = AddDataTableCols(dt);
            dt = AddDataTableRows(filesArray, dt);
            var userInputFileName = SpreadsheetNameInput();
            var fileExtension = spreadsheet.GetFileExtension();
            var newFilePath = Path.Combine(folderPath, userInputFileName + fileExtension);
            if (CheckIfFileExists(newFilePath) && !ShouldOverwriteFileName())
                _ = CreatFileNameCopy(folderPath, userInputFileName, fileExtension);
            spreadsheet.Create(dt, folderPath, userInputFileName);
        }

        /// <summary>
        /// Add columns to the DataTable object
        /// </summary>
        /// <param name="dataTable"></param>
        /// <returns></returns>
        public static System.Data.DataTable AddDataTableCols(System.Data.DataTable dataTable)
        {
            dataTable.Columns.Add("File Name", typeof(string));
            bool addNewCol = true;
            while (addNewCol)
            {
                Console.Write("Do you want to add another column? (Y/N): ");
                string moreColumns = Convert.ToString(Console.ReadLine());
                if (moreColumns.ToUpper().Equals("Y"))
                {
                    Console.Write("Type the name of the column: ");
                    string newCol = Console.ReadLine();
                    dataTable.Columns.Add(newCol, typeof(string));
                    Console.SetCursorPosition(0, Console.CursorTop - 1);
                    Console.Write(new string(' ', Console.WindowWidth));
                }
                else
                    addNewCol = false;
                Console.SetCursorPosition(0, Console.CursorTop - 1);
                Console.Write(new string(' ', Console.WindowWidth));
                Console.SetCursorPosition(0, Console.CursorTop);
            }
            return dataTable;
        }

        /// <summary>
        /// Add rows with the files array data to the DataTable object
        /// </summary>
        /// <param name="filesArray"></param>
        /// <param name="dataTable"></param>
        /// <returns></returns>
        public static System.Data.DataTable AddDataTableRows(string[] filesArray, System.Data.DataTable dataTable)
        {
            int filesArrayLength = filesArray.Length;
            int filesCount = 1;
            
            // Hide cursor
            Console.CursorVisible = false;
            // Loading message
            string loadingText = "Loading...";
            Console.Write(loadingText);
            Console.Write("{0}", filesCount.ToString("D" + filesArrayLength.ToString().Length));
            Console.Write(string.Format("/{0}", filesArrayLength));
            foreach (string file in filesArray)
            {
                if (filesCount > 1)
                {
                    var currentLineCursor = Console.CursorTop;
                    Console.SetCursorPosition(loadingText.Length, currentLineCursor);
                    Console.Write("{0}", filesCount.ToString("D" + filesArrayLength.ToString().Length));
                }
                filesCount++;
                dataTable.Rows.Add(Path.GetFileName(file));
            }
            // Show Cursor
            Console.CursorVisible = true;
            Console.WriteLine();
            return dataTable;
        }

        /// <summary>
        /// Asks user to insert a name for the spreadsheet file
        /// </summary>
        /// <returns></returns>
        public static string SpreadsheetNameInput ()
        {
            Console.Write("Insert the name of the spreadsheet file: ");
            string fileName = Console.ReadLine();
            while (CheckFileNameInvalidChars(fileName))
            {
                Console.WriteLine("Please, do not select a name with invalid characters for the spreadsheet file");
                Console.Write("Insert a new name for the spreadsheet file: ");
                fileName = Console.ReadLine();
            }
            return fileName;
        }

        /// <summary>
        /// Asks user If the file name should be overwritten when the file already exists
        /// </summary>
        /// <returns></returns>
        public static bool ShouldOverwriteFileName()
        {
            bool overwriteFile = true;
            Console.WriteLine("There's already a spreadsheet file with the given name");
            Console.Write("Do you want to overwrite the file? (Y/N): ");
            string overwriteFileName = Convert.ToString(Console.ReadLine());
            if (overwriteFileName.ToUpper().Equals("N") || overwriteFileName.ToUpper().Equals("NO"))
                overwriteFile = false;
            return overwriteFile;
        }

        /// <summary>
        /// Creates a new name when the file already exists
        /// </summary>
        /// <param name="folderPath"></param>
        /// <param name="pFileName"></param>
        /// <param name="pFILE_EXTENSION"></param>
        /// <returns></returns>
        public static string CreatFileNameCopy (string folderPath, string fileName, string fileExtension)
        {
            string copyWatermark = "_copy_";
            int copyNumber = 0;
            string newFileName, newFilePath;
            bool copyExists;
            do
            {
                copyNumber++;
                newFileName = $"{fileName}{copyWatermark}{copyNumber}";
                newFilePath = Path.Combine(folderPath, newFileName + fileExtension);
                copyExists = CheckIfFileExists(newFilePath);
            } while (copyExists);
            return newFilePath;
        }

        /// <summary>
        /// Checks If file name contains any invalid chars
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool CheckFileNameInvalidChars(string fileName)
        {
            return fileName.Any(Path.GetInvalidFileNameChars().Contains);
        }

        /// <summary>
        /// Checks If a file exists
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool CheckIfFileExists(string filePath)
        {
            return File.Exists(filePath);
        }
    }
}
