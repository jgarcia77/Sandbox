namespace Sandbox.OpenXML
{
    using System;
    using System.IO;

    public class Program
    {
        public static void Main(string[] args)
        {
            //CreateHyperlinkExample(Environment.CurrentDirectory);
            DStarExport(Environment.CurrentDirectory);
        }

        private static void CreateHyperlinkExample(string directory)
        {
            var path = Path.Combine(directory, "HyperlinkExample.xlsx");

            HyperlinkExample.Create(path);
        }

        private static void DStarExport(string directory)
        {
            var export = new GeneratedClass();

            export.CreatePackage(Path.Combine(directory, "DStarExport.xlsx"));
        }
    }
}
