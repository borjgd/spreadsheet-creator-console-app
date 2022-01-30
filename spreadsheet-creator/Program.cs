using System;

namespace spreadsheet_creator
{
    class Program
    {
        static void Main(string[] args)
        {
            string tmpPath = Environment.GetFolderPath(Environment.SpecialFolder.MyMusic);
            Console.WriteLine(tmpPath);
            Console.WriteLine("Press any key to finish");
            Console.ReadLine();
        }
    }
}
