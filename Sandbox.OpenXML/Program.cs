namespace Sandbox.OpenXML
{
    using System;
    using System.IO;

    public class Program
    {
        public static void Main(string[] args)
        {
            //CreateHyperlinkExample(Environment.CurrentDirectory);
            //DStarExport(Environment.CurrentDirectory);

            GetColumnName();

            Console.Read();
        }

        private static void GetColumnName()
        {
            var helper = new ExcelHelper();

            for (var index = 0; index <= (25 * 2); index++)
            {
                Console.WriteLine(helper.GetColumnName(index));
            }
        }

        private static void CreateHyperlinkExample(string directory)
        {
            var path = Path.Combine(directory, "HyperlinkExample.xlsx");

            HyperlinkExample.Create(path);
        }

        private static void DStarExport(string directory)
        {
            var export = new RevealExport();

            export.CreatePackage(Path.Combine(directory, "DStarExport.xlsx"));
        }
    }
}
