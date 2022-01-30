using System;
using System.IO;

namespace spreadsheet_creator
{
    class Program
    {
        static void Main(string[] args)
        {
            string tmpPath = Environment.GetFolderPath(Environment.SpecialFolder.MyMusic);
            Console.WriteLine(tmpPath);
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
                    Console.WriteLine(pathExistingFiles.Length + " files found");
                }
                else
                {
                    Console.WriteLine("No files found in this directory");
                }
            }
        }
    }
}
