namespace ExcelToXML;

internal class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
      //  ReadExcelDataHelpers ExcelHelper = new ReadExcelDataHelpers();
      //  ExcelHelper.ReadExcelData();
        IEnumerable<string> xlnames = ReadExcelDataHelpers.GetColumnValues("SampleTestData.xlsx", "A");
        ReadXMLDataHelpers XMLHelper = new ReadXMLDataHelpers();
        IEnumerable<string> xmlnames=XMLHelper.ReadXMLData();

        // Find elements in firstList that are not in secondList
        var notPresent = xlnames.Except(xmlnames).ToList();

        Console.WriteLine("Names not listed in the XML file:");
        foreach (var item in notPresent)
        {
            Console.WriteLine(item);
        }

    }
}
