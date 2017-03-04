namespace Sandbox.OpenXML
{
    using System;
    using System.IO;

    public class Program
    {
        public static void Main(string[] args)
        {
            CreateHyperlinkExample(Environment.CurrentDirectory);
        }

        private static void CreateHyperlinkExample(string directory)
        {
            var path = Path.Combine(directory, "HyperlinkExample.xlsx");

            HyperlinkExample.Create(path);
        }
    }
}
