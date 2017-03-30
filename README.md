# ExcelParser
Object Oriented Excel Parser
EPPlus for the basic Parsing

using `ExcelProperty` attribute, you can map the excel into your own object

    class Foo
    {
        [ExcelProperty("Bar")]
        public string Bar { get; private set; }

        [ExcelProperty("FooBar")]
        public string FooBar { get; private set; }
    }

a simple example:

    var excel = OExcel.Create(
                @"excel.xlsx");
    TempObj[] objs = excel.Worksheets[0].ReadAs<TempObj>(DataFlow.FirstColumnAsHeader);

you can read the excel either with "first-column-as-header" or "first-row-as-header"

