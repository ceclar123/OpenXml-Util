using OpenXml_Excel;

namespace OpenXml_Demo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("1 -> " + ExcelHelper.GetColName(1));
            Console.WriteLine("26 -> " + ExcelHelper.GetColName(26));
            Console.WriteLine("27 -> " + ExcelHelper.GetColName(27));
            Console.WriteLine("676 -> " + ExcelHelper.GetColName(676));            
            Console.WriteLine("677 -> " + ExcelHelper.GetColName(677));
            Console.WriteLine("702 -> " + ExcelHelper.GetColName(702));
            Console.WriteLine("703 -> " + ExcelHelper.GetColName(703));
            Console.WriteLine("723 -> " + ExcelHelper.GetColName(723));


            Console.WriteLine("Hello, World!");
        }
    }
}