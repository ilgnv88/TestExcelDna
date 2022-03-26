using ExcelDna.Integration;

namespace TestExcelDna
{
    public static class Functions
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string HelloDna(string name)
        {
            return "Hello " + name;
        }
    }
}
