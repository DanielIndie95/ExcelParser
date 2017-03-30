using System;

namespace ExcelParser
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelProperty : Attribute
    {
        public string ExcelHeader { get; set; }

        public ExcelProperty(string excelHeader)
        {
            ExcelHeader = excelHeader;
        }
    }
}
