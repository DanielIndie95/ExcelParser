using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelParser.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excel = OExcel.Create(
                @".xlsx");
            TempObj[] objs = excel.Worksheets[0].ReadAs<TempObj>(DataFlow.FirstColumnAsHeader);
            int a = 5;
        }
    }

    class TempObj
    {
        [ExcelProperty("Notes")]
        public string Notes { get; private set; }

        [ExcelProperty("English")]
        public string English { get; private set; }
    }
}
