using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ExcelParser
{
    public class OExcel
    {
        public OWorkSheet[] Worksheets { get; private set; }

        public OWorkSheet this[string name] => Worksheets.First(sheet => sheet.Name == name);

        public OExcel(string fileName)
        {
            Load(fileName);
        }

        public OExcel()
        {

        }

        public void Load(string fileName)
        {
            FileInfo file = new FileInfo(fileName);
            using (var excel = new ExcelPackage(file))
            {
                Worksheets = new OWorkSheet[excel.Workbook.Worksheets.Count];
                int worsheetIndex = 0;
                foreach (ExcelWorksheet workbookWorksheet in excel.Workbook.Worksheets)
                {
                    Worksheets[worsheetIndex] = new OWorkSheet(workbookWorksheet);
                }
            }
        }

        public static OExcel Create(string fileName)
        {
            return new OExcel(fileName);
        }
    }
}
