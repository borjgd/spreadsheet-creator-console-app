using DocumentFormat.OpenXml.Bibliography;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace spreadsheet_creator
{
    public class Spreadsheet
    {
        private readonly int rowIndex = 1;
        private readonly int columnIndex = 1;
        public const string FILE_EXTENSION = ".xlsx";
        public Spreadsheet() { }

        /// <summary>
        /// Creates a spreadsheet with SpreadSheetLight library
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="filesPath"></param>
        /// <param name="fileName"></param>
        public void Create(DataTable dataTable, string filesPath, string fileName)
        {
            SLDocument oSLDocument = new();
            Console.WriteLine("Importing data table to spreadsheet...");
            oSLDocument.ImportDataTable(rowIndex, columnIndex, dataTable, true);
            Console.Write("Saving file...");
            var filePath = Path.Combine(filesPath, fileName + FILE_EXTENSION);
            oSLDocument.SaveAs(filePath);
        }

        /// <summary>
        /// Gets the file extension
        /// </summary>
        /// <returns></returns>
        public string GetFileExtension() { return FILE_EXTENSION;}
    }
}
