namespace Sandbox.OpenXML
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    public static class HyperlinkExample
    {
        private static SpreadsheetDocument document;
        private static WorkbookPart workbookPart;
        private static Sheets sheets;

        public static void Create(string path)
        {
            document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);

            InitWorkbookPart();

            AddWorksheet();

            document.Close();
        }

        private static void InitWorkbookPart()
        {
            workbookPart = document.AddWorkbookPart();

            workbookPart.Workbook = new Workbook();

            sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
        }

        private static void AddWorksheet()
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            var worksheet = new Worksheet();

            var sheetData = new SheetData();

            FillSheetData(sheetData);

            worksheet.Append(sheetData);

            worksheetPart.Worksheet = worksheet;

            worksheetPart.Worksheet.Save();

            var sheet = new Sheet
            {
                Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = (uint)1,
                Name = "HyperlinkExample"
            };

            sheets.Append(sheet);

            document.WorkbookPart.Workbook.Save();
        }

        private static void FillSheetData(SheetData sheetData)
        {
            var row = new Row();

            var cell = new Cell
            {
                DataType = CellValues.String,
                CellValue = new CellValue("This is example of adding hyperlinks to a worksheet.")
            };

            row.Append(cell);

            sheetData.Append(row);
        }
    }
}
