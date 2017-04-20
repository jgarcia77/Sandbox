using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using Cs = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Sandbox.OpenXML
{
    public class ScatterChartProxy
    {
        // Creates a SpreadsheetDocument.
        public void CreatePackage(string filePath)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(SpreadsheetDocument document)
        {
            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPart1Content(worksheetPart1);

            DrawingsPart drawingsPart1 = worksheetPart1.AddNewPart<DrawingsPart>("rId2");
            GenerateDrawingsPart1Content(drawingsPart1);

            ChartPart chartPart1 = drawingsPart1.AddNewPart<ChartPart>("rId1");
            GenerateChartPart1Content(chartPart1);

            SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId5");
            GenerateSharedStringTablePart1Content(sharedStringTablePart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId4");
            GenerateWorkbookStylesPart1Content(workbookStylesPart1);
        }

        // Generates content of workbookPart1.
        private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x15" } };
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            workbook1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            workbook1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "6", LowestEdited = "6", BuildVersion = "14420" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties() { DefaultThemeVersion = (UInt32Value)153222U };

            AlternateContent alternateContent1 = new AlternateContent();
            alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "x15" };

            X15ac.AbsolutePath absolutePath1 = new X15ac.AbsolutePath() { Url = "C:\\Users\\josueg\\Documents\\Projects\\Magenic\\Deloitte\\STAR Modernization - Magenic Documents\\" };
            absolutePath1.AddNamespaceDeclaration("x15ac", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac");

            alternateContentChoice1.Append(absolutePath1);

            alternateContent1.Append(alternateContentChoice1);

            BookViews bookViews1 = new BookViews();
            WorkbookView workbookView1 = new WorkbookView() { XWindow = 0, YWindow = 0, WindowWidth = (UInt32Value)19200U, WindowHeight = (UInt32Value)7310U };

            bookViews1.Append(workbookView1);

            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Scatter Matrix", SheetId = (UInt32Value)3U, Id = "rId1" };

            sheets1.Append(sheet1);

            DefinedNames definedNames1 = new DefinedNames();
            DefinedName definedName1 = new DefinedName() { Name = "Multivariate_Data", LocalSheetId = (UInt32Value)0U };
            definedName1.Text = "\'Scatter Matrix\'!$A$1:$A$47";

            definedNames1.Append(definedName1);
            CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)152511U, CalculationOnSave = false };

            WorkbookExtensionList workbookExtensionList1 = new WorkbookExtensionList();

            WorkbookExtension workbookExtension1 = new WorkbookExtension() { Uri = "{140A7094-0E35-4892-8432-C4D2E57EDEB5}" };
            workbookExtension1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.WorkbookProperties workbookProperties2 = new X15.WorkbookProperties() { ChartTrackingReferenceBase = true };

            workbookExtension1.Append(workbookProperties2);

            workbookExtensionList1.Append(workbookExtension1);

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(alternateContent1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets1);
            workbook1.Append(definedNames1);
            workbook1.Append(calculationProperties1);
            workbook1.Append(workbookExtensionList1);

            workbookPart1.Workbook = workbook1;
        }

        // Generates content of worksheetPart1.
        private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:C47" };

            SheetViews sheetViews1 = new SheetViews();
            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultColumnWidth = 5.453125D, DefaultRowHeight = 14.5D, DyDescent = 0.35D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 5.54296875D, BestFit = true, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 5.6328125D, BestFit = true, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 6.90625D, BestFit = true, CustomWidth = true };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);

            SheetData sheetData1 = new SheetData();

            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, StyleIndex = (UInt32Value)1U, CustomFormat = true, DyDescent = 0.35D };

            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "0";

            cell1.Append(cellValue1);

            Cell cell2 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = "1";

            cell2.Append(cellValue2);

            row1.Append(cell1);
            row1.Append(cell2);

            Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell3 = new Cell() { CellReference = "A2" };
            CellValue cellValue3 = new CellValue();
            cellValue3.Text = "1";

            cell3.Append(cellValue3);

            Cell cell4 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue4 = new CellValue();
            cellValue4.Text = "0.43214092140921412";

            cell4.Append(cellValue4);

            Cell cell5 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue5 = new CellValue();
            cellValue5.Text = "2";

            cell5.Append(cellValue5);

            row2.Append(cell3);
            row2.Append(cell4);
            row2.Append(cell5);

            Row row3 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell6 = new Cell() { CellReference = "A3" };
            CellValue cellValue6 = new CellValue();
            cellValue6.Text = "2";

            cell6.Append(cellValue6);

            Cell cell7 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue7 = new CellValue();
            cellValue7.Text = "0.19528455284552845";

            cell7.Append(cellValue7);
            Cell cell8 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)4U };

            row3.Append(cell6);
            row3.Append(cell7);
            row3.Append(cell8);

            Row row4 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell9 = new Cell() { CellReference = "A4" };
            CellValue cellValue8 = new CellValue();
            cellValue8.Text = "3";

            cell9.Append(cellValue8);

            Cell cell10 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue9 = new CellValue();
            cellValue9.Text = "0.16422764227642275";

            cell10.Append(cellValue9);
            Cell cell11 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)4U };

            row4.Append(cell9);
            row4.Append(cell10);
            row4.Append(cell11);

            Row row5 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell12 = new Cell() { CellReference = "A5" };
            CellValue cellValue10 = new CellValue();
            cellValue10.Text = "4";

            cell12.Append(cellValue10);

            Cell cell13 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue11 = new CellValue();
            cellValue11.Text = "0.40249322493224932";

            cell13.Append(cellValue11);
            Cell cell14 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)4U };

            row5.Append(cell12);
            row5.Append(cell13);
            row5.Append(cell14);

            Row row6 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell15 = new Cell() { CellReference = "A6" };
            CellValue cellValue12 = new CellValue();
            cellValue12.Text = "5";

            cell15.Append(cellValue12);

            Cell cell16 = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue13 = new CellValue();
            cellValue13.Text = "0.15804878048780488";

            cell16.Append(cellValue13);
            Cell cell17 = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)4U };

            row6.Append(cell15);
            row6.Append(cell16);
            row6.Append(cell17);

            Row row7 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell18 = new Cell() { CellReference = "A7" };
            CellValue cellValue14 = new CellValue();
            cellValue14.Text = "6";

            cell18.Append(cellValue14);

            Cell cell19 = new Cell() { CellReference = "B7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue15 = new CellValue();
            cellValue15.Text = "0.23994579945799457";

            cell19.Append(cellValue15);
            Cell cell20 = new Cell() { CellReference = "C7", StyleIndex = (UInt32Value)4U };

            row7.Append(cell18);
            row7.Append(cell19);
            row7.Append(cell20);

            Row row8 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell21 = new Cell() { CellReference = "A8" };
            CellValue cellValue16 = new CellValue();
            cellValue16.Text = "7";

            cell21.Append(cellValue16);

            Cell cell22 = new Cell() { CellReference = "B8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue17 = new CellValue();
            cellValue17.Text = "0.40135501355013548";

            cell22.Append(cellValue17);
            Cell cell23 = new Cell() { CellReference = "C8", StyleIndex = (UInt32Value)4U };

            row8.Append(cell21);
            row8.Append(cell22);
            row8.Append(cell23);

            Row row9 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell24 = new Cell() { CellReference = "A9" };
            CellValue cellValue18 = new CellValue();
            cellValue18.Text = "8";

            cell24.Append(cellValue18);

            Cell cell25 = new Cell() { CellReference = "B9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue19 = new CellValue();
            cellValue19.Text = "0.26617886178861788";

            cell25.Append(cellValue19);
            Cell cell26 = new Cell() { CellReference = "C9", StyleIndex = (UInt32Value)4U };

            row9.Append(cell24);
            row9.Append(cell25);
            row9.Append(cell26);

            Row row10 = new Row() { RowIndex = (UInt32Value)10U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell27 = new Cell() { CellReference = "A10" };
            CellValue cellValue20 = new CellValue();
            cellValue20.Text = "9";

            cell27.Append(cellValue20);

            Cell cell28 = new Cell() { CellReference = "B10", StyleIndex = (UInt32Value)2U };
            CellValue cellValue21 = new CellValue();
            cellValue21.Text = "0.19479674796747967";

            cell28.Append(cellValue21);
            Cell cell29 = new Cell() { CellReference = "C10", StyleIndex = (UInt32Value)4U };

            row10.Append(cell27);
            row10.Append(cell28);
            row10.Append(cell29);

            Row row11 = new Row() { RowIndex = (UInt32Value)11U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell30 = new Cell() { CellReference = "A11" };
            CellValue cellValue22 = new CellValue();
            cellValue22.Text = "10";

            cell30.Append(cellValue22);

            Cell cell31 = new Cell() { CellReference = "B11", StyleIndex = (UInt32Value)2U };
            CellValue cellValue23 = new CellValue();
            cellValue23.Text = "0.47306233062330622";

            cell31.Append(cellValue23);
            Cell cell32 = new Cell() { CellReference = "C11", StyleIndex = (UInt32Value)4U };

            row11.Append(cell30);
            row11.Append(cell31);
            row11.Append(cell32);

            Row row12 = new Row() { RowIndex = (UInt32Value)12U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell33 = new Cell() { CellReference = "A12" };
            CellValue cellValue24 = new CellValue();
            cellValue24.Text = "11";

            cell33.Append(cellValue24);

            Cell cell34 = new Cell() { CellReference = "B12", StyleIndex = (UInt32Value)2U };
            CellValue cellValue25 = new CellValue();
            cellValue25.Text = "0.22639566395663957";

            cell34.Append(cellValue25);
            Cell cell35 = new Cell() { CellReference = "C12", StyleIndex = (UInt32Value)4U };

            row12.Append(cell33);
            row12.Append(cell34);
            row12.Append(cell35);

            Row row13 = new Row() { RowIndex = (UInt32Value)13U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell36 = new Cell() { CellReference = "A13" };
            CellValue cellValue26 = new CellValue();
            cellValue26.Text = "12";

            cell36.Append(cellValue26);

            Cell cell37 = new Cell() { CellReference = "B13", StyleIndex = (UInt32Value)2U };
            CellValue cellValue27 = new CellValue();
            cellValue27.Text = "0";

            cell37.Append(cellValue27);
            Cell cell38 = new Cell() { CellReference = "C13", StyleIndex = (UInt32Value)4U };

            row13.Append(cell36);
            row13.Append(cell37);
            row13.Append(cell38);

            Row row14 = new Row() { RowIndex = (UInt32Value)14U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell39 = new Cell() { CellReference = "A14" };
            CellValue cellValue28 = new CellValue();
            cellValue28.Text = "13";

            cell39.Append(cellValue28);

            Cell cell40 = new Cell() { CellReference = "B14", StyleIndex = (UInt32Value)2U };
            CellValue cellValue29 = new CellValue();
            cellValue29.Text = "0.59701897018970185";

            cell40.Append(cellValue29);
            Cell cell41 = new Cell() { CellReference = "C14", StyleIndex = (UInt32Value)4U };

            row14.Append(cell39);
            row14.Append(cell40);
            row14.Append(cell41);

            Row row15 = new Row() { RowIndex = (UInt32Value)15U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell42 = new Cell() { CellReference = "A15" };
            CellValue cellValue30 = new CellValue();
            cellValue30.Text = "14";

            cell42.Append(cellValue30);

            Cell cell43 = new Cell() { CellReference = "B15", StyleIndex = (UInt32Value)2U };
            CellValue cellValue31 = new CellValue();
            cellValue31.Text = "0.33880758807588074";

            cell43.Append(cellValue31);
            Cell cell44 = new Cell() { CellReference = "C15", StyleIndex = (UInt32Value)4U };

            row15.Append(cell42);
            row15.Append(cell43);
            row15.Append(cell44);

            Row row16 = new Row() { RowIndex = (UInt32Value)16U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell45 = new Cell() { CellReference = "A16" };
            CellValue cellValue32 = new CellValue();
            cellValue32.Text = "15";

            cell45.Append(cellValue32);

            Cell cell46 = new Cell() { CellReference = "B16", StyleIndex = (UInt32Value)2U };
            CellValue cellValue33 = new CellValue();
            cellValue33.Text = "0.38195121951219513";

            cell46.Append(cellValue33);
            Cell cell47 = new Cell() { CellReference = "C16", StyleIndex = (UInt32Value)4U };

            row16.Append(cell45);
            row16.Append(cell46);
            row16.Append(cell47);

            Row row17 = new Row() { RowIndex = (UInt32Value)17U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell48 = new Cell() { CellReference = "A17" };
            CellValue cellValue34 = new CellValue();
            cellValue34.Text = "16";

            cell48.Append(cellValue34);

            Cell cell49 = new Cell() { CellReference = "B17", StyleIndex = (UInt32Value)2U };
            CellValue cellValue35 = new CellValue();
            cellValue35.Text = "0.57642276422764227";

            cell49.Append(cellValue35);
            Cell cell50 = new Cell() { CellReference = "C17", StyleIndex = (UInt32Value)4U };

            row17.Append(cell48);
            row17.Append(cell49);
            row17.Append(cell50);

            Row row18 = new Row() { RowIndex = (UInt32Value)18U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell51 = new Cell() { CellReference = "A18" };
            CellValue cellValue36 = new CellValue();
            cellValue36.Text = "17";

            cell51.Append(cellValue36);

            Cell cell52 = new Cell() { CellReference = "B18", StyleIndex = (UInt32Value)2U };
            CellValue cellValue37 = new CellValue();
            cellValue37.Text = "0.33983739837398375";

            cell52.Append(cellValue37);
            Cell cell53 = new Cell() { CellReference = "C18", StyleIndex = (UInt32Value)4U };

            row18.Append(cell51);
            row18.Append(cell52);
            row18.Append(cell53);

            Row row19 = new Row() { RowIndex = (UInt32Value)19U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell54 = new Cell() { CellReference = "A19" };
            CellValue cellValue38 = new CellValue();
            cellValue38.Text = "18";

            cell54.Append(cellValue38);

            Cell cell55 = new Cell() { CellReference = "B19", StyleIndex = (UInt32Value)2U };
            CellValue cellValue39 = new CellValue();
            cellValue39.Text = "0.39002710027100274";

            cell55.Append(cellValue39);
            Cell cell56 = new Cell() { CellReference = "C19", StyleIndex = (UInt32Value)4U };

            row19.Append(cell54);
            row19.Append(cell55);
            row19.Append(cell56);

            Row row20 = new Row() { RowIndex = (UInt32Value)20U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell57 = new Cell() { CellReference = "A20" };
            CellValue cellValue40 = new CellValue();
            cellValue40.Text = "19";

            cell57.Append(cellValue40);

            Cell cell58 = new Cell() { CellReference = "B20", StyleIndex = (UInt32Value)2U };
            CellValue cellValue41 = new CellValue();
            cellValue41.Text = "0.57604336043360438";

            cell58.Append(cellValue41);
            Cell cell59 = new Cell() { CellReference = "C20", StyleIndex = (UInt32Value)4U };

            row20.Append(cell57);
            row20.Append(cell58);
            row20.Append(cell59);

            Row row21 = new Row() { RowIndex = (UInt32Value)21U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell60 = new Cell() { CellReference = "A21" };
            CellValue cellValue42 = new CellValue();
            cellValue42.Text = "20";

            cell60.Append(cellValue42);

            Cell cell61 = new Cell() { CellReference = "B21", StyleIndex = (UInt32Value)2U };
            CellValue cellValue43 = new CellValue();
            cellValue43.Text = "0.43425474254742547";

            cell61.Append(cellValue43);
            Cell cell62 = new Cell() { CellReference = "C21", StyleIndex = (UInt32Value)4U };

            row21.Append(cell60);
            row21.Append(cell61);
            row21.Append(cell62);

            Row row22 = new Row() { RowIndex = (UInt32Value)22U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell63 = new Cell() { CellReference = "A22" };
            CellValue cellValue44 = new CellValue();
            cellValue44.Text = "21";

            cell63.Append(cellValue44);

            Cell cell64 = new Cell() { CellReference = "B22", StyleIndex = (UInt32Value)2U };
            CellValue cellValue45 = new CellValue();
            cellValue45.Text = "0.39203252032520325";

            cell64.Append(cellValue45);
            Cell cell65 = new Cell() { CellReference = "C22", StyleIndex = (UInt32Value)4U };

            row22.Append(cell63);
            row22.Append(cell64);
            row22.Append(cell65);

            Row row23 = new Row() { RowIndex = (UInt32Value)23U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell66 = new Cell() { CellReference = "A23" };
            CellValue cellValue46 = new CellValue();
            cellValue46.Text = "22";

            cell66.Append(cellValue46);

            Cell cell67 = new Cell() { CellReference = "B23", StyleIndex = (UInt32Value)2U };
            CellValue cellValue47 = new CellValue();
            cellValue47.Text = "0.7223306233062331";

            cell67.Append(cellValue47);
            Cell cell68 = new Cell() { CellReference = "C23", StyleIndex = (UInt32Value)4U };

            row23.Append(cell66);
            row23.Append(cell67);
            row23.Append(cell68);

            Row row24 = new Row() { RowIndex = (UInt32Value)24U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell69 = new Cell() { CellReference = "A24" };
            CellValue cellValue48 = new CellValue();
            cellValue48.Text = "23";

            cell69.Append(cellValue48);

            Cell cell70 = new Cell() { CellReference = "B24", StyleIndex = (UInt32Value)2U };
            CellValue cellValue49 = new CellValue();
            cellValue49.Text = "0.42807588075880759";

            cell70.Append(cellValue49);
            Cell cell71 = new Cell() { CellReference = "C24", StyleIndex = (UInt32Value)4U };

            row24.Append(cell69);
            row24.Append(cell70);
            row24.Append(cell71);

            Row row25 = new Row() { RowIndex = (UInt32Value)25U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell72 = new Cell() { CellReference = "A25" };
            CellValue cellValue50 = new CellValue();
            cellValue50.Text = "24";

            cell72.Append(cellValue50);

            Cell cell73 = new Cell() { CellReference = "B25", StyleIndex = (UInt32Value)2U };
            CellValue cellValue51 = new CellValue();
            cellValue51.Text = "0.2088888888888889";

            cell73.Append(cellValue51);
            Cell cell74 = new Cell() { CellReference = "C25", StyleIndex = (UInt32Value)4U };

            row25.Append(cell72);
            row25.Append(cell73);
            row25.Append(cell74);

            Row row26 = new Row() { RowIndex = (UInt32Value)26U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell75 = new Cell() { CellReference = "A26" };
            CellValue cellValue52 = new CellValue();
            cellValue52.Text = "25";

            cell75.Append(cellValue52);

            Cell cell76 = new Cell() { CellReference = "B26", StyleIndex = (UInt32Value)2U };
            CellValue cellValue53 = new CellValue();
            cellValue53.Text = "0.95241192411924114";

            cell76.Append(cellValue53);
            Cell cell77 = new Cell() { CellReference = "C26", StyleIndex = (UInt32Value)4U };

            row26.Append(cell75);
            row26.Append(cell76);
            row26.Append(cell77);

            Row row27 = new Row() { RowIndex = (UInt32Value)27U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell78 = new Cell() { CellReference = "A27" };
            CellValue cellValue54 = new CellValue();
            cellValue54.Text = "26";

            cell78.Append(cellValue54);

            Cell cell79 = new Cell() { CellReference = "B27", StyleIndex = (UInt32Value)2U };
            CellValue cellValue55 = new CellValue();
            cellValue55.Text = "0.63934959349593501";

            cell79.Append(cellValue55);
            Cell cell80 = new Cell() { CellReference = "C27", StyleIndex = (UInt32Value)4U };

            row27.Append(cell78);
            row27.Append(cell79);
            row27.Append(cell80);

            Row row28 = new Row() { RowIndex = (UInt32Value)28U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell81 = new Cell() { CellReference = "A28" };
            CellValue cellValue56 = new CellValue();
            cellValue56.Text = "27";

            cell81.Append(cellValue56);

            Cell cell82 = new Cell() { CellReference = "B28", StyleIndex = (UInt32Value)2U };
            CellValue cellValue57 = new CellValue();
            cellValue57.Text = "0.61409214092140918";

            cell82.Append(cellValue57);
            Cell cell83 = new Cell() { CellReference = "C28", StyleIndex = (UInt32Value)4U };

            row28.Append(cell81);
            row28.Append(cell82);
            row28.Append(cell83);

            Row row29 = new Row() { RowIndex = (UInt32Value)29U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell84 = new Cell() { CellReference = "A29" };
            CellValue cellValue58 = new CellValue();
            cellValue58.Text = "28";

            cell84.Append(cellValue58);

            Cell cell85 = new Cell() { CellReference = "B29", StyleIndex = (UInt32Value)2U };
            CellValue cellValue59 = new CellValue();
            cellValue59.Text = "0.75062330623306228";

            cell85.Append(cellValue59);
            Cell cell86 = new Cell() { CellReference = "C29", StyleIndex = (UInt32Value)4U };

            row29.Append(cell84);
            row29.Append(cell85);
            row29.Append(cell86);

            Row row30 = new Row() { RowIndex = (UInt32Value)30U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell87 = new Cell() { CellReference = "A30" };
            CellValue cellValue60 = new CellValue();
            cellValue60.Text = "29";

            cell87.Append(cellValue60);

            Cell cell88 = new Cell() { CellReference = "B30", StyleIndex = (UInt32Value)2U };
            CellValue cellValue61 = new CellValue();
            cellValue61.Text = "0.51100271002710029";

            cell88.Append(cellValue61);
            Cell cell89 = new Cell() { CellReference = "C30", StyleIndex = (UInt32Value)4U };

            row30.Append(cell87);
            row30.Append(cell88);
            row30.Append(cell89);

            Row row31 = new Row() { RowIndex = (UInt32Value)31U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell90 = new Cell() { CellReference = "A31" };
            CellValue cellValue62 = new CellValue();
            cellValue62.Text = "30";

            cell90.Append(cellValue62);

            Cell cell91 = new Cell() { CellReference = "B31", StyleIndex = (UInt32Value)2U };
            CellValue cellValue63 = new CellValue();
            cellValue63.Text = "0.5095934959349594";

            cell91.Append(cellValue63);
            Cell cell92 = new Cell() { CellReference = "C31", StyleIndex = (UInt32Value)4U };

            row31.Append(cell90);
            row31.Append(cell91);
            row31.Append(cell92);

            Row row32 = new Row() { RowIndex = (UInt32Value)32U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell93 = new Cell() { CellReference = "A32" };
            CellValue cellValue64 = new CellValue();
            cellValue64.Text = "31";

            cell93.Append(cellValue64);

            Cell cell94 = new Cell() { CellReference = "B32", StyleIndex = (UInt32Value)2U };
            CellValue cellValue65 = new CellValue();
            cellValue65.Text = "0.68574525745257453";

            cell94.Append(cellValue65);
            Cell cell95 = new Cell() { CellReference = "C32", StyleIndex = (UInt32Value)4U };

            row32.Append(cell93);
            row32.Append(cell94);
            row32.Append(cell95);

            Row row33 = new Row() { RowIndex = (UInt32Value)33U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell96 = new Cell() { CellReference = "A33" };
            CellValue cellValue66 = new CellValue();
            cellValue66.Text = "32";

            cell96.Append(cellValue66);

            Cell cell97 = new Cell() { CellReference = "B33", StyleIndex = (UInt32Value)2U };
            CellValue cellValue67 = new CellValue();
            cellValue67.Text = "0.48899728997289971";

            cell97.Append(cellValue67);
            Cell cell98 = new Cell() { CellReference = "C33", StyleIndex = (UInt32Value)4U };

            row33.Append(cell96);
            row33.Append(cell97);
            row33.Append(cell98);

            Row row34 = new Row() { RowIndex = (UInt32Value)34U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell99 = new Cell() { CellReference = "A34" };
            CellValue cellValue68 = new CellValue();
            cellValue68.Text = "33";

            cell99.Append(cellValue68);

            Cell cell100 = new Cell() { CellReference = "B34", StyleIndex = (UInt32Value)2U };
            CellValue cellValue69 = new CellValue();
            cellValue69.Text = "0.42780487804878048";

            cell100.Append(cellValue69);
            Cell cell101 = new Cell() { CellReference = "C34", StyleIndex = (UInt32Value)4U };

            row34.Append(cell99);
            row34.Append(cell100);
            row34.Append(cell101);

            Row row35 = new Row() { RowIndex = (UInt32Value)35U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell102 = new Cell() { CellReference = "A35" };
            CellValue cellValue70 = new CellValue();
            cellValue70.Text = "34";

            cell102.Append(cellValue70);

            Cell cell103 = new Cell() { CellReference = "B35", StyleIndex = (UInt32Value)2U };
            CellValue cellValue71 = new CellValue();
            cellValue71.Text = "0.75766937669376688";

            cell103.Append(cellValue71);
            Cell cell104 = new Cell() { CellReference = "C35", StyleIndex = (UInt32Value)4U };

            row35.Append(cell102);
            row35.Append(cell103);
            row35.Append(cell104);

            Row row36 = new Row() { RowIndex = (UInt32Value)36U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell105 = new Cell() { CellReference = "A36" };
            CellValue cellValue72 = new CellValue();
            cellValue72.Text = "35";

            cell105.Append(cellValue72);

            Cell cell106 = new Cell() { CellReference = "B36", StyleIndex = (UInt32Value)2U };
            CellValue cellValue73 = new CellValue();
            cellValue73.Text = "0.4897560975609756";

            cell106.Append(cellValue73);
            Cell cell107 = new Cell() { CellReference = "C36", StyleIndex = (UInt32Value)4U };

            row36.Append(cell105);
            row36.Append(cell106);
            row36.Append(cell107);

            Row row37 = new Row() { RowIndex = (UInt32Value)37U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell108 = new Cell() { CellReference = "A37" };
            CellValue cellValue74 = new CellValue();
            cellValue74.Text = "36";

            cell108.Append(cellValue74);

            Cell cell109 = new Cell() { CellReference = "B37", StyleIndex = (UInt32Value)2U };
            CellValue cellValue75 = new CellValue();
            cellValue75.Text = "0.16043360433604337";

            cell109.Append(cellValue75);
            Cell cell110 = new Cell() { CellReference = "C37", StyleIndex = (UInt32Value)4U };

            row37.Append(cell108);
            row37.Append(cell109);
            row37.Append(cell110);

            Row row38 = new Row() { RowIndex = (UInt32Value)38U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell111 = new Cell() { CellReference = "A38", StyleIndex = (UInt32Value)1U };
            CellValue cellValue76 = new CellValue();
            cellValue76.Text = "37";

            cell111.Append(cellValue76);

            Cell cell112 = new Cell() { CellReference = "B38", StyleIndex = (UInt32Value)3U };
            CellValue cellValue77 = new CellValue();
            cellValue77.Text = "1";

            cell112.Append(cellValue77);

            Cell cell113 = new Cell() { CellReference = "C38", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue78 = new CellValue();
            cellValue78.Text = "3";

            cell113.Append(cellValue78);

            row38.Append(cell111);
            row38.Append(cell112);
            row38.Append(cell113);

            Row row39 = new Row() { RowIndex = (UInt32Value)39U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell114 = new Cell() { CellReference = "A39", StyleIndex = (UInt32Value)1U };
            CellValue cellValue79 = new CellValue();
            cellValue79.Text = "38";

            cell114.Append(cellValue79);

            Cell cell115 = new Cell() { CellReference = "B39", StyleIndex = (UInt32Value)3U };
            CellValue cellValue80 = new CellValue();
            cellValue80.Text = "0.57018970189701901";

            cell115.Append(cellValue80);
            Cell cell116 = new Cell() { CellReference = "C39", StyleIndex = (UInt32Value)5U };

            row39.Append(cell114);
            row39.Append(cell115);
            row39.Append(cell116);

            Row row40 = new Row() { RowIndex = (UInt32Value)40U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell117 = new Cell() { CellReference = "A40", StyleIndex = (UInt32Value)1U };
            CellValue cellValue81 = new CellValue();
            cellValue81.Text = "39";

            cell117.Append(cellValue81);

            Cell cell118 = new Cell() { CellReference = "B40", StyleIndex = (UInt32Value)3U };
            CellValue cellValue82 = new CellValue();
            cellValue82.Text = "0.59355013550135505";

            cell118.Append(cellValue82);
            Cell cell119 = new Cell() { CellReference = "C40", StyleIndex = (UInt32Value)5U };

            row40.Append(cell117);
            row40.Append(cell118);
            row40.Append(cell119);

            Row row41 = new Row() { RowIndex = (UInt32Value)41U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell120 = new Cell() { CellReference = "A41", StyleIndex = (UInt32Value)1U };
            CellValue cellValue83 = new CellValue();
            cellValue83.Text = "40";

            cell120.Append(cellValue83);

            Cell cell121 = new Cell() { CellReference = "B41", StyleIndex = (UInt32Value)3U };
            CellValue cellValue84 = new CellValue();
            cellValue84.Text = "0.78899728997289975";

            cell121.Append(cellValue84);
            Cell cell122 = new Cell() { CellReference = "C41", StyleIndex = (UInt32Value)5U };

            row41.Append(cell120);
            row41.Append(cell121);
            row41.Append(cell122);

            Row row42 = new Row() { RowIndex = (UInt32Value)42U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell123 = new Cell() { CellReference = "A42", StyleIndex = (UInt32Value)1U };
            CellValue cellValue85 = new CellValue();
            cellValue85.Text = "41";

            cell123.Append(cellValue85);

            Cell cell124 = new Cell() { CellReference = "B42", StyleIndex = (UInt32Value)3U };
            CellValue cellValue86 = new CellValue();
            cellValue86.Text = "0.51056910569105696";

            cell124.Append(cellValue86);
            Cell cell125 = new Cell() { CellReference = "C42", StyleIndex = (UInt32Value)5U };

            row42.Append(cell123);
            row42.Append(cell124);
            row42.Append(cell125);

            Row row43 = new Row() { RowIndex = (UInt32Value)43U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell126 = new Cell() { CellReference = "A43", StyleIndex = (UInt32Value)1U };
            CellValue cellValue87 = new CellValue();
            cellValue87.Text = "42";

            cell126.Append(cellValue87);

            Cell cell127 = new Cell() { CellReference = "B43", StyleIndex = (UInt32Value)3U };
            CellValue cellValue88 = new CellValue();
            cellValue88.Text = "0.56336043360433607";

            cell127.Append(cellValue88);
            Cell cell128 = new Cell() { CellReference = "C43", StyleIndex = (UInt32Value)5U };

            row43.Append(cell126);
            row43.Append(cell127);
            row43.Append(cell128);

            Row row44 = new Row() { RowIndex = (UInt32Value)44U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell129 = new Cell() { CellReference = "A44", StyleIndex = (UInt32Value)1U };
            CellValue cellValue89 = new CellValue();
            cellValue89.Text = "43";

            cell129.Append(cellValue89);

            Cell cell130 = new Cell() { CellReference = "B44", StyleIndex = (UInt32Value)3U };
            CellValue cellValue90 = new CellValue();
            cellValue90.Text = "0.77653116531165312";

            cell130.Append(cellValue90);
            Cell cell131 = new Cell() { CellReference = "C44", StyleIndex = (UInt32Value)5U };

            row44.Append(cell129);
            row44.Append(cell130);
            row44.Append(cell131);

            Row row45 = new Row() { RowIndex = (UInt32Value)45U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell132 = new Cell() { CellReference = "A45", StyleIndex = (UInt32Value)1U };
            CellValue cellValue91 = new CellValue();
            cellValue91.Text = "44";

            cell132.Append(cellValue91);

            Cell cell133 = new Cell() { CellReference = "B45", StyleIndex = (UInt32Value)3U };
            CellValue cellValue92 = new CellValue();
            cellValue92.Text = "0.55116531165311655";

            cell133.Append(cellValue92);
            Cell cell134 = new Cell() { CellReference = "C45", StyleIndex = (UInt32Value)5U };

            row45.Append(cell132);
            row45.Append(cell133);
            row45.Append(cell134);

            Row row46 = new Row() { RowIndex = (UInt32Value)46U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell135 = new Cell() { CellReference = "A46", StyleIndex = (UInt32Value)1U };
            CellValue cellValue93 = new CellValue();
            cellValue93.Text = "45";

            cell135.Append(cellValue93);

            Cell cell136 = new Cell() { CellReference = "B46", StyleIndex = (UInt32Value)3U };
            CellValue cellValue94 = new CellValue();
            cellValue94.Text = "0.48298102981029811";

            cell136.Append(cellValue94);
            Cell cell137 = new Cell() { CellReference = "C46", StyleIndex = (UInt32Value)5U };

            row46.Append(cell135);
            row46.Append(cell136);
            row46.Append(cell137);

            Row row47 = new Row() { RowIndex = (UInt32Value)47U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.35D };

            Cell cell138 = new Cell() { CellReference = "A47", StyleIndex = (UInt32Value)1U };
            CellValue cellValue95 = new CellValue();
            cellValue95.Text = "46";

            cell138.Append(cellValue95);

            Cell cell139 = new Cell() { CellReference = "B47", StyleIndex = (UInt32Value)3U };
            CellValue cellValue96 = new CellValue();
            cellValue96.Text = "0.89311653116531164";

            cell139.Append(cellValue96);
            Cell cell140 = new Cell() { CellReference = "C47", StyleIndex = (UInt32Value)5U };

            row47.Append(cell138);
            row47.Append(cell139);
            row47.Append(cell140);

            sheetData1.Append(row1);
            sheetData1.Append(row2);
            sheetData1.Append(row3);
            sheetData1.Append(row4);
            sheetData1.Append(row5);
            sheetData1.Append(row6);
            sheetData1.Append(row7);
            sheetData1.Append(row8);
            sheetData1.Append(row9);
            sheetData1.Append(row10);
            sheetData1.Append(row11);
            sheetData1.Append(row12);
            sheetData1.Append(row13);
            sheetData1.Append(row14);
            sheetData1.Append(row15);
            sheetData1.Append(row16);
            sheetData1.Append(row17);
            sheetData1.Append(row18);
            sheetData1.Append(row19);
            sheetData1.Append(row20);
            sheetData1.Append(row21);
            sheetData1.Append(row22);
            sheetData1.Append(row23);
            sheetData1.Append(row24);
            sheetData1.Append(row25);
            sheetData1.Append(row26);
            sheetData1.Append(row27);
            sheetData1.Append(row28);
            sheetData1.Append(row29);
            sheetData1.Append(row30);
            sheetData1.Append(row31);
            sheetData1.Append(row32);
            sheetData1.Append(row33);
            sheetData1.Append(row34);
            sheetData1.Append(row35);
            sheetData1.Append(row36);
            sheetData1.Append(row37);
            sheetData1.Append(row38);
            sheetData1.Append(row39);
            sheetData1.Append(row40);
            sheetData1.Append(row41);
            sheetData1.Append(row42);
            sheetData1.Append(row43);
            sheetData1.Append(row44);
            sheetData1.Append(row45);
            sheetData1.Append(row46);
            sheetData1.Append(row47);

            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)2U };
            MergeCell mergeCell1 = new MergeCell() { Reference = "C2:C37" };
            MergeCell mergeCell2 = new MergeCell() { Reference = "C38:C47" };

            mergeCells1.Append(mergeCell1);
            mergeCells1.Append(mergeCell2);
            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup() { Orientation = OrientationValues.Portrait, Id = "rId1" };
            Drawing drawing1 = new Drawing() { Id = "rId2" };

            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(mergeCells1);
            worksheet1.Append(pageMargins1);
            worksheet1.Append(pageSetup1);
            worksheet1.Append(drawing1);

            worksheetPart1.Worksheet = worksheet1;
        }

        // Generates content of drawingsPart1.
        private void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1)
        {
            Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor();

            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            columnId1.Text = "4";
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            columnOffset1.Text = "371475";
            Xdr.RowId rowId1 = new Xdr.RowId();
            rowId1.Text = "2";
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            rowOffset1.Text = "6350";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            columnId2.Text = "9";
            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            columnOffset2.Text = "342900";
            Xdr.RowId rowId2 = new Xdr.RowId();
            rowId2.Text = "10";
            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            rowOffset2.Text = "171450";

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            Xdr.GraphicFrame graphicFrame1 = new Xdr.GraphicFrame() { Macro = "" };

            Xdr.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = new Xdr.NonVisualGraphicFrameProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Chart 1" };
            Xdr.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Xdr.NonVisualGraphicFrameDrawingProperties();

            nonVisualGraphicFrameProperties1.Append(nonVisualDrawingProperties1);
            nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameDrawingProperties1);

            Xdr.Transform transform1 = new Xdr.Transform();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 0L, Cy = 0L };

            transform1.Append(offset1);
            transform1.Append(extents1);

            A.Graphic graphic1 = new A.Graphic();

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            C.ChartReference chartReference1 = new C.ChartReference() { Id = "rId1" };
            chartReference1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartReference1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            graphicData1.Append(chartReference1);

            graphic1.Append(graphicData1);

            graphicFrame1.Append(nonVisualGraphicFrameProperties1);
            graphicFrame1.Append(transform1);
            graphicFrame1.Append(graphic1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(graphicFrame1);
            twoCellAnchor1.Append(clientData1);

            worksheetDrawing1.Append(twoCellAnchor1);

            drawingsPart1.WorksheetDrawing = worksheetDrawing1;
        }

        // Generates content of chartPart1.
        private void GenerateChartPart1Content(ChartPart chartPart1)
        {
            C.ChartSpace chartSpace1 = new C.ChartSpace();
            chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            C.Date1904 date19041 = new C.Date1904() { Val = false };
            C.EditingLanguage editingLanguage1 = new C.EditingLanguage() { Val = "en-US" };
            C.RoundedCorners roundedCorners1 = new C.RoundedCorners() { Val = false };

            AlternateContent alternateContent2 = new AlternateContent();
            alternateContent2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice2 = new AlternateContentChoice() { Requires = "c14" };
            alternateContentChoice2.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
            C14.Style style1 = new C14.Style() { Val = 102 };

            alternateContentChoice2.Append(style1);

            AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();
            C.Style style2 = new C.Style() { Val = 2 };

            alternateContentFallback1.Append(style2);

            alternateContent2.Append(alternateContentChoice2);
            alternateContent2.Append(alternateContentFallback1);

            C.Chart chart1 = new C.Chart();
            C.AutoTitleDeleted autoTitleDeleted1 = new C.AutoTitleDeleted() { Val = true };

            C.PlotArea plotArea1 = new C.PlotArea();

            C.Layout layout1 = new C.Layout();

            C.ManualLayout manualLayout1 = new C.ManualLayout();
            C.LayoutTarget layoutTarget1 = new C.LayoutTarget() { Val = C.LayoutTargetValues.Inner };
            C.LeftMode leftMode1 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
            C.TopMode topMode1 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
            C.Left left1 = new C.Left() { Val = 0.23737532808398951D };
            C.Top top1 = new C.Top() { Val = 0.16074185848720129D };
            C.Width width1 = new C.Width() { Val = 0.7118469052127977D };
            C.Height height1 = new C.Height() { Val = 0.78870202200334727D };

            manualLayout1.Append(layoutTarget1);
            manualLayout1.Append(leftMode1);
            manualLayout1.Append(topMode1);
            manualLayout1.Append(left1);
            manualLayout1.Append(top1);
            manualLayout1.Append(width1);
            manualLayout1.Append(height1);

            layout1.Append(manualLayout1);

            C.ScatterChart scatterChart1 = new C.ScatterChart();
            C.ScatterStyle scatterStyle1 = new C.ScatterStyle() { Val = C.ScatterStyleValues.LineMarker };
            C.VaryColors varyColors1 = new C.VaryColors() { Val = false };

            C.ScatterChartSeries scatterChartSeries1 = new C.ScatterChartSeries();
            C.Index index1 = new C.Index() { Val = (UInt32Value)1U };
            C.Order order1 = new C.Order() { Val = (UInt32Value)0U };

            C.ChartShapeProperties chartShapeProperties1 = new C.ChartShapeProperties();

            A.Outline outline4 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Round };
            A.NoFill noFill1 = new A.NoFill();
            A.Round round1 = new A.Round();

            outline4.Append(noFill1);
            outline4.Append(round1);
            A.EffectList effectList4 = new A.EffectList();

            chartShapeProperties1.Append(outline4);
            chartShapeProperties1.Append(effectList4);

            C.Marker marker1 = new C.Marker();
            C.Symbol symbol1 = new C.Symbol() { Val = C.MarkerStyleValues.Circle };
            C.Size size1 = new C.Size() { Val = 5 };

            C.ChartShapeProperties chartShapeProperties2 = new C.ChartShapeProperties();
            A.NoFill noFill2 = new A.NoFill();

            A.Outline outline5 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill7 = new A.SolidFill();
            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            solidFill7.Append(schemeColor16);

            outline5.Append(solidFill7);
            A.EffectList effectList5 = new A.EffectList();

            chartShapeProperties2.Append(noFill2);
            chartShapeProperties2.Append(outline5);
            chartShapeProperties2.Append(effectList5);

            marker1.Append(symbol1);
            marker1.Append(size1);
            marker1.Append(chartShapeProperties2);

            C.XValues xValues1 = new C.XValues();

            C.NumberReference numberReference1 = new C.NumberReference();
            C.Formula formula1 = new C.Formula();
            formula1.Text = "\'Scatter Matrix\'!$B$2:$B$37";

            C.NumberingCache numberingCache1 = new C.NumberingCache();
            C.FormatCode formatCode1 = new C.FormatCode();
            formatCode1.Text = "#,##0.000";
            C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)36U };

            C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue1 = new C.NumericValue();
            numericValue1.Text = "0.43214092140921412";

            numericPoint1.Append(numericValue1);

            C.NumericPoint numericPoint2 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue2 = new C.NumericValue();
            numericValue2.Text = "0.19528455284552845";

            numericPoint2.Append(numericValue2);

            C.NumericPoint numericPoint3 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue3 = new C.NumericValue();
            numericValue3.Text = "0.16422764227642275";

            numericPoint3.Append(numericValue3);

            C.NumericPoint numericPoint4 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue4 = new C.NumericValue();
            numericValue4.Text = "0.40249322493224932";

            numericPoint4.Append(numericValue4);

            C.NumericPoint numericPoint5 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue5 = new C.NumericValue();
            numericValue5.Text = "0.15804878048780488";

            numericPoint5.Append(numericValue5);

            C.NumericPoint numericPoint6 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue6 = new C.NumericValue();
            numericValue6.Text = "0.23994579945799457";

            numericPoint6.Append(numericValue6);

            C.NumericPoint numericPoint7 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue7 = new C.NumericValue();
            numericValue7.Text = "0.40135501355013548";

            numericPoint7.Append(numericValue7);

            C.NumericPoint numericPoint8 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue8 = new C.NumericValue();
            numericValue8.Text = "0.26617886178861788";

            numericPoint8.Append(numericValue8);

            C.NumericPoint numericPoint9 = new C.NumericPoint() { Index = (UInt32Value)8U };
            C.NumericValue numericValue9 = new C.NumericValue();
            numericValue9.Text = "0.19479674796747967";

            numericPoint9.Append(numericValue9);

            C.NumericPoint numericPoint10 = new C.NumericPoint() { Index = (UInt32Value)9U };
            C.NumericValue numericValue10 = new C.NumericValue();
            numericValue10.Text = "0.47306233062330622";

            numericPoint10.Append(numericValue10);

            C.NumericPoint numericPoint11 = new C.NumericPoint() { Index = (UInt32Value)10U };
            C.NumericValue numericValue11 = new C.NumericValue();
            numericValue11.Text = "0.22639566395663957";

            numericPoint11.Append(numericValue11);

            C.NumericPoint numericPoint12 = new C.NumericPoint() { Index = (UInt32Value)11U };
            C.NumericValue numericValue12 = new C.NumericValue();
            numericValue12.Text = "0";

            numericPoint12.Append(numericValue12);

            C.NumericPoint numericPoint13 = new C.NumericPoint() { Index = (UInt32Value)12U };
            C.NumericValue numericValue13 = new C.NumericValue();
            numericValue13.Text = "0.59701897018970185";

            numericPoint13.Append(numericValue13);

            C.NumericPoint numericPoint14 = new C.NumericPoint() { Index = (UInt32Value)13U };
            C.NumericValue numericValue14 = new C.NumericValue();
            numericValue14.Text = "0.33880758807588074";

            numericPoint14.Append(numericValue14);

            C.NumericPoint numericPoint15 = new C.NumericPoint() { Index = (UInt32Value)14U };
            C.NumericValue numericValue15 = new C.NumericValue();
            numericValue15.Text = "0.38195121951219513";

            numericPoint15.Append(numericValue15);

            C.NumericPoint numericPoint16 = new C.NumericPoint() { Index = (UInt32Value)15U };
            C.NumericValue numericValue16 = new C.NumericValue();
            numericValue16.Text = "0.57642276422764227";

            numericPoint16.Append(numericValue16);

            C.NumericPoint numericPoint17 = new C.NumericPoint() { Index = (UInt32Value)16U };
            C.NumericValue numericValue17 = new C.NumericValue();
            numericValue17.Text = "0.33983739837398375";

            numericPoint17.Append(numericValue17);

            C.NumericPoint numericPoint18 = new C.NumericPoint() { Index = (UInt32Value)17U };
            C.NumericValue numericValue18 = new C.NumericValue();
            numericValue18.Text = "0.39002710027100274";

            numericPoint18.Append(numericValue18);

            C.NumericPoint numericPoint19 = new C.NumericPoint() { Index = (UInt32Value)18U };
            C.NumericValue numericValue19 = new C.NumericValue();
            numericValue19.Text = "0.57604336043360438";

            numericPoint19.Append(numericValue19);

            C.NumericPoint numericPoint20 = new C.NumericPoint() { Index = (UInt32Value)19U };
            C.NumericValue numericValue20 = new C.NumericValue();
            numericValue20.Text = "0.43425474254742547";

            numericPoint20.Append(numericValue20);

            C.NumericPoint numericPoint21 = new C.NumericPoint() { Index = (UInt32Value)20U };
            C.NumericValue numericValue21 = new C.NumericValue();
            numericValue21.Text = "0.39203252032520325";

            numericPoint21.Append(numericValue21);

            C.NumericPoint numericPoint22 = new C.NumericPoint() { Index = (UInt32Value)21U };
            C.NumericValue numericValue22 = new C.NumericValue();
            numericValue22.Text = "0.7223306233062331";

            numericPoint22.Append(numericValue22);

            C.NumericPoint numericPoint23 = new C.NumericPoint() { Index = (UInt32Value)22U };
            C.NumericValue numericValue23 = new C.NumericValue();
            numericValue23.Text = "0.42807588075880759";

            numericPoint23.Append(numericValue23);

            C.NumericPoint numericPoint24 = new C.NumericPoint() { Index = (UInt32Value)23U };
            C.NumericValue numericValue24 = new C.NumericValue();
            numericValue24.Text = "0.2088888888888889";

            numericPoint24.Append(numericValue24);

            C.NumericPoint numericPoint25 = new C.NumericPoint() { Index = (UInt32Value)24U };
            C.NumericValue numericValue25 = new C.NumericValue();
            numericValue25.Text = "0.95241192411924114";

            numericPoint25.Append(numericValue25);

            C.NumericPoint numericPoint26 = new C.NumericPoint() { Index = (UInt32Value)25U };
            C.NumericValue numericValue26 = new C.NumericValue();
            numericValue26.Text = "0.63934959349593501";

            numericPoint26.Append(numericValue26);

            C.NumericPoint numericPoint27 = new C.NumericPoint() { Index = (UInt32Value)26U };
            C.NumericValue numericValue27 = new C.NumericValue();
            numericValue27.Text = "0.61409214092140918";

            numericPoint27.Append(numericValue27);

            C.NumericPoint numericPoint28 = new C.NumericPoint() { Index = (UInt32Value)27U };
            C.NumericValue numericValue28 = new C.NumericValue();
            numericValue28.Text = "0.75062330623306228";

            numericPoint28.Append(numericValue28);

            C.NumericPoint numericPoint29 = new C.NumericPoint() { Index = (UInt32Value)28U };
            C.NumericValue numericValue29 = new C.NumericValue();
            numericValue29.Text = "0.51100271002710029";

            numericPoint29.Append(numericValue29);

            C.NumericPoint numericPoint30 = new C.NumericPoint() { Index = (UInt32Value)29U };
            C.NumericValue numericValue30 = new C.NumericValue();
            numericValue30.Text = "0.5095934959349594";

            numericPoint30.Append(numericValue30);

            C.NumericPoint numericPoint31 = new C.NumericPoint() { Index = (UInt32Value)30U };
            C.NumericValue numericValue31 = new C.NumericValue();
            numericValue31.Text = "0.68574525745257453";

            numericPoint31.Append(numericValue31);

            C.NumericPoint numericPoint32 = new C.NumericPoint() { Index = (UInt32Value)31U };
            C.NumericValue numericValue32 = new C.NumericValue();
            numericValue32.Text = "0.48899728997289971";

            numericPoint32.Append(numericValue32);

            C.NumericPoint numericPoint33 = new C.NumericPoint() { Index = (UInt32Value)32U };
            C.NumericValue numericValue33 = new C.NumericValue();
            numericValue33.Text = "0.42780487804878048";

            numericPoint33.Append(numericValue33);

            C.NumericPoint numericPoint34 = new C.NumericPoint() { Index = (UInt32Value)33U };
            C.NumericValue numericValue34 = new C.NumericValue();
            numericValue34.Text = "0.75766937669376688";

            numericPoint34.Append(numericValue34);

            C.NumericPoint numericPoint35 = new C.NumericPoint() { Index = (UInt32Value)34U };
            C.NumericValue numericValue35 = new C.NumericValue();
            numericValue35.Text = "0.4897560975609756";

            numericPoint35.Append(numericValue35);

            C.NumericPoint numericPoint36 = new C.NumericPoint() { Index = (UInt32Value)35U };
            C.NumericValue numericValue36 = new C.NumericValue();
            numericValue36.Text = "0.16043360433604337";

            numericPoint36.Append(numericValue36);

            numberingCache1.Append(formatCode1);
            numberingCache1.Append(pointCount1);
            numberingCache1.Append(numericPoint1);
            numberingCache1.Append(numericPoint2);
            numberingCache1.Append(numericPoint3);
            numberingCache1.Append(numericPoint4);
            numberingCache1.Append(numericPoint5);
            numberingCache1.Append(numericPoint6);
            numberingCache1.Append(numericPoint7);
            numberingCache1.Append(numericPoint8);
            numberingCache1.Append(numericPoint9);
            numberingCache1.Append(numericPoint10);
            numberingCache1.Append(numericPoint11);
            numberingCache1.Append(numericPoint12);
            numberingCache1.Append(numericPoint13);
            numberingCache1.Append(numericPoint14);
            numberingCache1.Append(numericPoint15);
            numberingCache1.Append(numericPoint16);
            numberingCache1.Append(numericPoint17);
            numberingCache1.Append(numericPoint18);
            numberingCache1.Append(numericPoint19);
            numberingCache1.Append(numericPoint20);
            numberingCache1.Append(numericPoint21);
            numberingCache1.Append(numericPoint22);
            numberingCache1.Append(numericPoint23);
            numberingCache1.Append(numericPoint24);
            numberingCache1.Append(numericPoint25);
            numberingCache1.Append(numericPoint26);
            numberingCache1.Append(numericPoint27);
            numberingCache1.Append(numericPoint28);
            numberingCache1.Append(numericPoint29);
            numberingCache1.Append(numericPoint30);
            numberingCache1.Append(numericPoint31);
            numberingCache1.Append(numericPoint32);
            numberingCache1.Append(numericPoint33);
            numberingCache1.Append(numericPoint34);
            numberingCache1.Append(numericPoint35);
            numberingCache1.Append(numericPoint36);

            numberReference1.Append(formula1);
            numberReference1.Append(numberingCache1);

            xValues1.Append(numberReference1);

            C.YValues yValues1 = new C.YValues();

            C.NumberReference numberReference2 = new C.NumberReference();
            C.Formula formula2 = new C.Formula();
            formula2.Text = "\'Scatter Matrix\'!$B$2:$B$37";

            C.NumberingCache numberingCache2 = new C.NumberingCache();
            C.FormatCode formatCode2 = new C.FormatCode();
            formatCode2.Text = "#,##0.000";
            C.PointCount pointCount2 = new C.PointCount() { Val = (UInt32Value)36U };

            C.NumericPoint numericPoint37 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue37 = new C.NumericValue();
            numericValue37.Text = "0.43214092140921412";

            numericPoint37.Append(numericValue37);

            C.NumericPoint numericPoint38 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue38 = new C.NumericValue();
            numericValue38.Text = "0.19528455284552845";

            numericPoint38.Append(numericValue38);

            C.NumericPoint numericPoint39 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue39 = new C.NumericValue();
            numericValue39.Text = "0.16422764227642275";

            numericPoint39.Append(numericValue39);

            C.NumericPoint numericPoint40 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue40 = new C.NumericValue();
            numericValue40.Text = "0.40249322493224932";

            numericPoint40.Append(numericValue40);

            C.NumericPoint numericPoint41 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue41 = new C.NumericValue();
            numericValue41.Text = "0.15804878048780488";

            numericPoint41.Append(numericValue41);

            C.NumericPoint numericPoint42 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue42 = new C.NumericValue();
            numericValue42.Text = "0.23994579945799457";

            numericPoint42.Append(numericValue42);

            C.NumericPoint numericPoint43 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue43 = new C.NumericValue();
            numericValue43.Text = "0.40135501355013548";

            numericPoint43.Append(numericValue43);

            C.NumericPoint numericPoint44 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue44 = new C.NumericValue();
            numericValue44.Text = "0.26617886178861788";

            numericPoint44.Append(numericValue44);

            C.NumericPoint numericPoint45 = new C.NumericPoint() { Index = (UInt32Value)8U };
            C.NumericValue numericValue45 = new C.NumericValue();
            numericValue45.Text = "0.19479674796747967";

            numericPoint45.Append(numericValue45);

            C.NumericPoint numericPoint46 = new C.NumericPoint() { Index = (UInt32Value)9U };
            C.NumericValue numericValue46 = new C.NumericValue();
            numericValue46.Text = "0.47306233062330622";

            numericPoint46.Append(numericValue46);

            C.NumericPoint numericPoint47 = new C.NumericPoint() { Index = (UInt32Value)10U };
            C.NumericValue numericValue47 = new C.NumericValue();
            numericValue47.Text = "0.22639566395663957";

            numericPoint47.Append(numericValue47);

            C.NumericPoint numericPoint48 = new C.NumericPoint() { Index = (UInt32Value)11U };
            C.NumericValue numericValue48 = new C.NumericValue();
            numericValue48.Text = "0";

            numericPoint48.Append(numericValue48);

            C.NumericPoint numericPoint49 = new C.NumericPoint() { Index = (UInt32Value)12U };
            C.NumericValue numericValue49 = new C.NumericValue();
            numericValue49.Text = "0.59701897018970185";

            numericPoint49.Append(numericValue49);

            C.NumericPoint numericPoint50 = new C.NumericPoint() { Index = (UInt32Value)13U };
            C.NumericValue numericValue50 = new C.NumericValue();
            numericValue50.Text = "0.33880758807588074";

            numericPoint50.Append(numericValue50);

            C.NumericPoint numericPoint51 = new C.NumericPoint() { Index = (UInt32Value)14U };
            C.NumericValue numericValue51 = new C.NumericValue();
            numericValue51.Text = "0.38195121951219513";

            numericPoint51.Append(numericValue51);

            C.NumericPoint numericPoint52 = new C.NumericPoint() { Index = (UInt32Value)15U };
            C.NumericValue numericValue52 = new C.NumericValue();
            numericValue52.Text = "0.57642276422764227";

            numericPoint52.Append(numericValue52);

            C.NumericPoint numericPoint53 = new C.NumericPoint() { Index = (UInt32Value)16U };
            C.NumericValue numericValue53 = new C.NumericValue();
            numericValue53.Text = "0.33983739837398375";

            numericPoint53.Append(numericValue53);

            C.NumericPoint numericPoint54 = new C.NumericPoint() { Index = (UInt32Value)17U };
            C.NumericValue numericValue54 = new C.NumericValue();
            numericValue54.Text = "0.39002710027100274";

            numericPoint54.Append(numericValue54);

            C.NumericPoint numericPoint55 = new C.NumericPoint() { Index = (UInt32Value)18U };
            C.NumericValue numericValue55 = new C.NumericValue();
            numericValue55.Text = "0.57604336043360438";

            numericPoint55.Append(numericValue55);

            C.NumericPoint numericPoint56 = new C.NumericPoint() { Index = (UInt32Value)19U };
            C.NumericValue numericValue56 = new C.NumericValue();
            numericValue56.Text = "0.43425474254742547";

            numericPoint56.Append(numericValue56);

            C.NumericPoint numericPoint57 = new C.NumericPoint() { Index = (UInt32Value)20U };
            C.NumericValue numericValue57 = new C.NumericValue();
            numericValue57.Text = "0.39203252032520325";

            numericPoint57.Append(numericValue57);

            C.NumericPoint numericPoint58 = new C.NumericPoint() { Index = (UInt32Value)21U };
            C.NumericValue numericValue58 = new C.NumericValue();
            numericValue58.Text = "0.7223306233062331";

            numericPoint58.Append(numericValue58);

            C.NumericPoint numericPoint59 = new C.NumericPoint() { Index = (UInt32Value)22U };
            C.NumericValue numericValue59 = new C.NumericValue();
            numericValue59.Text = "0.42807588075880759";

            numericPoint59.Append(numericValue59);

            C.NumericPoint numericPoint60 = new C.NumericPoint() { Index = (UInt32Value)23U };
            C.NumericValue numericValue60 = new C.NumericValue();
            numericValue60.Text = "0.2088888888888889";

            numericPoint60.Append(numericValue60);

            C.NumericPoint numericPoint61 = new C.NumericPoint() { Index = (UInt32Value)24U };
            C.NumericValue numericValue61 = new C.NumericValue();
            numericValue61.Text = "0.95241192411924114";

            numericPoint61.Append(numericValue61);

            C.NumericPoint numericPoint62 = new C.NumericPoint() { Index = (UInt32Value)25U };
            C.NumericValue numericValue62 = new C.NumericValue();
            numericValue62.Text = "0.63934959349593501";

            numericPoint62.Append(numericValue62);

            C.NumericPoint numericPoint63 = new C.NumericPoint() { Index = (UInt32Value)26U };
            C.NumericValue numericValue63 = new C.NumericValue();
            numericValue63.Text = "0.61409214092140918";

            numericPoint63.Append(numericValue63);

            C.NumericPoint numericPoint64 = new C.NumericPoint() { Index = (UInt32Value)27U };
            C.NumericValue numericValue64 = new C.NumericValue();
            numericValue64.Text = "0.75062330623306228";

            numericPoint64.Append(numericValue64);

            C.NumericPoint numericPoint65 = new C.NumericPoint() { Index = (UInt32Value)28U };
            C.NumericValue numericValue65 = new C.NumericValue();
            numericValue65.Text = "0.51100271002710029";

            numericPoint65.Append(numericValue65);

            C.NumericPoint numericPoint66 = new C.NumericPoint() { Index = (UInt32Value)29U };
            C.NumericValue numericValue66 = new C.NumericValue();
            numericValue66.Text = "0.5095934959349594";

            numericPoint66.Append(numericValue66);

            C.NumericPoint numericPoint67 = new C.NumericPoint() { Index = (UInt32Value)30U };
            C.NumericValue numericValue67 = new C.NumericValue();
            numericValue67.Text = "0.68574525745257453";

            numericPoint67.Append(numericValue67);

            C.NumericPoint numericPoint68 = new C.NumericPoint() { Index = (UInt32Value)31U };
            C.NumericValue numericValue68 = new C.NumericValue();
            numericValue68.Text = "0.48899728997289971";

            numericPoint68.Append(numericValue68);

            C.NumericPoint numericPoint69 = new C.NumericPoint() { Index = (UInt32Value)32U };
            C.NumericValue numericValue69 = new C.NumericValue();
            numericValue69.Text = "0.42780487804878048";

            numericPoint69.Append(numericValue69);

            C.NumericPoint numericPoint70 = new C.NumericPoint() { Index = (UInt32Value)33U };
            C.NumericValue numericValue70 = new C.NumericValue();
            numericValue70.Text = "0.75766937669376688";

            numericPoint70.Append(numericValue70);

            C.NumericPoint numericPoint71 = new C.NumericPoint() { Index = (UInt32Value)34U };
            C.NumericValue numericValue71 = new C.NumericValue();
            numericValue71.Text = "0.4897560975609756";

            numericPoint71.Append(numericValue71);

            C.NumericPoint numericPoint72 = new C.NumericPoint() { Index = (UInt32Value)35U };
            C.NumericValue numericValue72 = new C.NumericValue();
            numericValue72.Text = "0.16043360433604337";

            numericPoint72.Append(numericValue72);

            numberingCache2.Append(formatCode2);
            numberingCache2.Append(pointCount2);
            numberingCache2.Append(numericPoint37);
            numberingCache2.Append(numericPoint38);
            numberingCache2.Append(numericPoint39);
            numberingCache2.Append(numericPoint40);
            numberingCache2.Append(numericPoint41);
            numberingCache2.Append(numericPoint42);
            numberingCache2.Append(numericPoint43);
            numberingCache2.Append(numericPoint44);
            numberingCache2.Append(numericPoint45);
            numberingCache2.Append(numericPoint46);
            numberingCache2.Append(numericPoint47);
            numberingCache2.Append(numericPoint48);
            numberingCache2.Append(numericPoint49);
            numberingCache2.Append(numericPoint50);
            numberingCache2.Append(numericPoint51);
            numberingCache2.Append(numericPoint52);
            numberingCache2.Append(numericPoint53);
            numberingCache2.Append(numericPoint54);
            numberingCache2.Append(numericPoint55);
            numberingCache2.Append(numericPoint56);
            numberingCache2.Append(numericPoint57);
            numberingCache2.Append(numericPoint58);
            numberingCache2.Append(numericPoint59);
            numberingCache2.Append(numericPoint60);
            numberingCache2.Append(numericPoint61);
            numberingCache2.Append(numericPoint62);
            numberingCache2.Append(numericPoint63);
            numberingCache2.Append(numericPoint64);
            numberingCache2.Append(numericPoint65);
            numberingCache2.Append(numericPoint66);
            numberingCache2.Append(numericPoint67);
            numberingCache2.Append(numericPoint68);
            numberingCache2.Append(numericPoint69);
            numberingCache2.Append(numericPoint70);
            numberingCache2.Append(numericPoint71);
            numberingCache2.Append(numericPoint72);

            numberReference2.Append(formula2);
            numberReference2.Append(numberingCache2);

            yValues1.Append(numberReference2);
            C.Smooth smooth1 = new C.Smooth() { Val = false };

            scatterChartSeries1.Append(index1);
            scatterChartSeries1.Append(order1);
            scatterChartSeries1.Append(chartShapeProperties1);
            scatterChartSeries1.Append(marker1);
            scatterChartSeries1.Append(xValues1);
            scatterChartSeries1.Append(yValues1);
            scatterChartSeries1.Append(smooth1);

            C.ScatterChartSeries scatterChartSeries2 = new C.ScatterChartSeries();
            C.Index index2 = new C.Index() { Val = (UInt32Value)0U };
            C.Order order2 = new C.Order() { Val = (UInt32Value)1U };

            C.ChartShapeProperties chartShapeProperties3 = new C.ChartShapeProperties();

            A.Outline outline6 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Round };
            A.NoFill noFill3 = new A.NoFill();
            A.Round round2 = new A.Round();

            outline6.Append(noFill3);
            outline6.Append(round2);
            A.EffectList effectList6 = new A.EffectList();

            chartShapeProperties3.Append(outline6);
            chartShapeProperties3.Append(effectList6);

            C.Marker marker2 = new C.Marker();
            C.Symbol symbol2 = new C.Symbol() { Val = C.MarkerStyleValues.X };
            C.Size size2 = new C.Size() { Val = 5 };

            C.ChartShapeProperties chartShapeProperties4 = new C.ChartShapeProperties();
            A.NoFill noFill4 = new A.NoFill();

            A.Outline outline7 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat };

            A.SolidFill solidFill8 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "FF0000" };

            solidFill8.Append(rgbColorModelHex12);
            A.LineJoinBevel lineJoinBevel1 = new A.LineJoinBevel();

            outline7.Append(solidFill8);
            outline7.Append(lineJoinBevel1);
            A.EffectList effectList7 = new A.EffectList();

            chartShapeProperties4.Append(noFill4);
            chartShapeProperties4.Append(outline7);
            chartShapeProperties4.Append(effectList7);

            marker2.Append(symbol2);
            marker2.Append(size2);
            marker2.Append(chartShapeProperties4);

            C.XValues xValues2 = new C.XValues();

            C.NumberReference numberReference3 = new C.NumberReference();
            C.Formula formula3 = new C.Formula();
            formula3.Text = "\'Scatter Matrix\'!$B$38:$B$47";

            C.NumberingCache numberingCache3 = new C.NumberingCache();
            C.FormatCode formatCode3 = new C.FormatCode();
            formatCode3.Text = "#,##0.000";
            C.PointCount pointCount3 = new C.PointCount() { Val = (UInt32Value)10U };

            C.NumericPoint numericPoint73 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue73 = new C.NumericValue();
            numericValue73.Text = "1";

            numericPoint73.Append(numericValue73);

            C.NumericPoint numericPoint74 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue74 = new C.NumericValue();
            numericValue74.Text = "0.57018970189701901";

            numericPoint74.Append(numericValue74);

            C.NumericPoint numericPoint75 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue75 = new C.NumericValue();
            numericValue75.Text = "0.59355013550135505";

            numericPoint75.Append(numericValue75);

            C.NumericPoint numericPoint76 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue76 = new C.NumericValue();
            numericValue76.Text = "0.78899728997289975";

            numericPoint76.Append(numericValue76);

            C.NumericPoint numericPoint77 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue77 = new C.NumericValue();
            numericValue77.Text = "0.51056910569105696";

            numericPoint77.Append(numericValue77);

            C.NumericPoint numericPoint78 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue78 = new C.NumericValue();
            numericValue78.Text = "0.56336043360433607";

            numericPoint78.Append(numericValue78);

            C.NumericPoint numericPoint79 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue79 = new C.NumericValue();
            numericValue79.Text = "0.77653116531165312";

            numericPoint79.Append(numericValue79);

            C.NumericPoint numericPoint80 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue80 = new C.NumericValue();
            numericValue80.Text = "0.55116531165311655";

            numericPoint80.Append(numericValue80);

            C.NumericPoint numericPoint81 = new C.NumericPoint() { Index = (UInt32Value)8U };
            C.NumericValue numericValue81 = new C.NumericValue();
            numericValue81.Text = "0.48298102981029811";

            numericPoint81.Append(numericValue81);

            C.NumericPoint numericPoint82 = new C.NumericPoint() { Index = (UInt32Value)9U };
            C.NumericValue numericValue82 = new C.NumericValue();
            numericValue82.Text = "0.89311653116531164";

            numericPoint82.Append(numericValue82);

            numberingCache3.Append(formatCode3);
            numberingCache3.Append(pointCount3);
            numberingCache3.Append(numericPoint73);
            numberingCache3.Append(numericPoint74);
            numberingCache3.Append(numericPoint75);
            numberingCache3.Append(numericPoint76);
            numberingCache3.Append(numericPoint77);
            numberingCache3.Append(numericPoint78);
            numberingCache3.Append(numericPoint79);
            numberingCache3.Append(numericPoint80);
            numberingCache3.Append(numericPoint81);
            numberingCache3.Append(numericPoint82);

            numberReference3.Append(formula3);
            numberReference3.Append(numberingCache3);

            xValues2.Append(numberReference3);

            C.YValues yValues2 = new C.YValues();

            C.NumberReference numberReference4 = new C.NumberReference();
            C.Formula formula4 = new C.Formula();
            formula4.Text = "\'Scatter Matrix\'!$B$38:$B$47";

            C.NumberingCache numberingCache4 = new C.NumberingCache();
            C.FormatCode formatCode4 = new C.FormatCode();
            formatCode4.Text = "#,##0.000";
            C.PointCount pointCount4 = new C.PointCount() { Val = (UInt32Value)10U };

            C.NumericPoint numericPoint83 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue83 = new C.NumericValue();
            numericValue83.Text = "1";

            numericPoint83.Append(numericValue83);

            C.NumericPoint numericPoint84 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue84 = new C.NumericValue();
            numericValue84.Text = "0.57018970189701901";

            numericPoint84.Append(numericValue84);

            C.NumericPoint numericPoint85 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue85 = new C.NumericValue();
            numericValue85.Text = "0.59355013550135505";

            numericPoint85.Append(numericValue85);

            C.NumericPoint numericPoint86 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue86 = new C.NumericValue();
            numericValue86.Text = "0.78899728997289975";

            numericPoint86.Append(numericValue86);

            C.NumericPoint numericPoint87 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue87 = new C.NumericValue();
            numericValue87.Text = "0.51056910569105696";

            numericPoint87.Append(numericValue87);

            C.NumericPoint numericPoint88 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue88 = new C.NumericValue();
            numericValue88.Text = "0.56336043360433607";

            numericPoint88.Append(numericValue88);

            C.NumericPoint numericPoint89 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue89 = new C.NumericValue();
            numericValue89.Text = "0.77653116531165312";

            numericPoint89.Append(numericValue89);

            C.NumericPoint numericPoint90 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue90 = new C.NumericValue();
            numericValue90.Text = "0.55116531165311655";

            numericPoint90.Append(numericValue90);

            C.NumericPoint numericPoint91 = new C.NumericPoint() { Index = (UInt32Value)8U };
            C.NumericValue numericValue91 = new C.NumericValue();
            numericValue91.Text = "0.48298102981029811";

            numericPoint91.Append(numericValue91);

            C.NumericPoint numericPoint92 = new C.NumericPoint() { Index = (UInt32Value)9U };
            C.NumericValue numericValue92 = new C.NumericValue();
            numericValue92.Text = "0.89311653116531164";

            numericPoint92.Append(numericValue92);

            numberingCache4.Append(formatCode4);
            numberingCache4.Append(pointCount4);
            numberingCache4.Append(numericPoint83);
            numberingCache4.Append(numericPoint84);
            numberingCache4.Append(numericPoint85);
            numberingCache4.Append(numericPoint86);
            numberingCache4.Append(numericPoint87);
            numberingCache4.Append(numericPoint88);
            numberingCache4.Append(numericPoint89);
            numberingCache4.Append(numericPoint90);
            numberingCache4.Append(numericPoint91);
            numberingCache4.Append(numericPoint92);

            numberReference4.Append(formula4);
            numberReference4.Append(numberingCache4);

            yValues2.Append(numberReference4);
            C.Smooth smooth2 = new C.Smooth() { Val = false };

            scatterChartSeries2.Append(index2);
            scatterChartSeries2.Append(order2);
            scatterChartSeries2.Append(chartShapeProperties3);
            scatterChartSeries2.Append(marker2);
            scatterChartSeries2.Append(xValues2);
            scatterChartSeries2.Append(yValues2);
            scatterChartSeries2.Append(smooth2);

            C.DataLabels dataLabels1 = new C.DataLabels();
            C.ShowLegendKey showLegendKey1 = new C.ShowLegendKey() { Val = false };
            C.ShowValue showValue1 = new C.ShowValue() { Val = false };
            C.ShowCategoryName showCategoryName1 = new C.ShowCategoryName() { Val = false };
            C.ShowSeriesName showSeriesName1 = new C.ShowSeriesName() { Val = false };
            C.ShowPercent showPercent1 = new C.ShowPercent() { Val = false };
            C.ShowBubbleSize showBubbleSize1 = new C.ShowBubbleSize() { Val = false };

            dataLabels1.Append(showLegendKey1);
            dataLabels1.Append(showValue1);
            dataLabels1.Append(showCategoryName1);
            dataLabels1.Append(showSeriesName1);
            dataLabels1.Append(showPercent1);
            dataLabels1.Append(showBubbleSize1);
            C.AxisId axisId1 = new C.AxisId() { Val = (UInt32Value)375001568U };
            C.AxisId axisId2 = new C.AxisId() { Val = (UInt32Value)375001960U };

            scatterChart1.Append(scatterStyle1);
            scatterChart1.Append(varyColors1);
            scatterChart1.Append(scatterChartSeries1);
            scatterChart1.Append(scatterChartSeries2);
            //scatterChart1.Append(dataLabels1);
            scatterChart1.Append(axisId1);
            scatterChart1.Append(axisId2);

            C.ValueAxis valueAxis1 = new C.ValueAxis();
            C.AxisId axisId3 = new C.AxisId() { Val = (UInt32Value)375001568U };

            C.Scaling scaling1 = new C.Scaling();
            C.Orientation orientation1 = new C.Orientation() { Val = C.OrientationValues.MinMax };
            C.MaxAxisValue maxAxisValue1 = new C.MaxAxisValue() { Val = 1D };

            scaling1.Append(orientation1);
            scaling1.Append(maxAxisValue1);
            C.Delete delete1 = new C.Delete() { Val = true };
            C.AxisPosition axisPosition1 = new C.AxisPosition() { Val = C.AxisPositionValues.Bottom };

            C.MajorGridlines majorGridlines1 = new C.MajorGridlines();

            C.ChartShapeProperties chartShapeProperties5 = new C.ChartShapeProperties();

            A.Outline outline8 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill9 = new A.SolidFill();

            A.SchemeColor schemeColor17 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation9 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor17.Append(luminanceModulation9);
            schemeColor17.Append(luminanceOffset1);

            solidFill9.Append(schemeColor17);
            A.Round round3 = new A.Round();

            outline8.Append(solidFill9);
            outline8.Append(round3);
            A.EffectList effectList8 = new A.EffectList();

            chartShapeProperties5.Append(outline8);
            chartShapeProperties5.Append(effectList8);

            majorGridlines1.Append(chartShapeProperties5);

            C.Title title1 = new C.Title();

            C.ChartText chartText1 = new C.ChartText();

            C.RichText richText1 = new C.RichText();
            A.BodyProperties bodyProperties1 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties() { FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill10 = new A.SolidFill();

            A.SchemeColor schemeColor18 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation10 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor18.Append(luminanceModulation10);
            schemeColor18.Append(luminanceOffset2);

            solidFill10.Append(schemeColor18);
            A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont3 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties1.Append(solidFill10);
            defaultRunProperties1.Append(latinFont3);
            defaultRunProperties1.Append(eastAsianFont3);
            defaultRunProperties1.Append(complexScriptFont3);

            paragraphProperties1.Append(defaultRunProperties1);

            A.Run run1 = new A.Run();
            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US", Bold = true };
            A.Text text1 = new A.Text();
            text1.Text = "Sales";

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            richText1.Append(bodyProperties1);
            richText1.Append(listStyle1);
            richText1.Append(paragraph1);

            chartText1.Append(richText1);

            C.Layout layout2 = new C.Layout();

            C.ManualLayout manualLayout2 = new C.ManualLayout();
            C.LeftMode leftMode2 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
            C.TopMode topMode2 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
            C.Left left2 = new C.Left() { Val = 0.48899786260894601D };
            C.Top top2 = new C.Top() { Val = 4.1314164997667976E-2D };

            manualLayout2.Append(leftMode2);
            manualLayout2.Append(topMode2);
            manualLayout2.Append(left2);
            manualLayout2.Append(top2);

            layout2.Append(manualLayout2);
            C.Overlay overlay1 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties6 = new C.ChartShapeProperties();
            A.NoFill noFill5 = new A.NoFill();

            A.Outline outline9 = new A.Outline();
            A.NoFill noFill6 = new A.NoFill();

            outline9.Append(noFill6);
            A.EffectList effectList9 = new A.EffectList();

            chartShapeProperties6.Append(noFill5);
            chartShapeProperties6.Append(outline9);
            chartShapeProperties6.Append(effectList9);

            C.TextProperties textProperties2 = new C.TextProperties();
            A.BodyProperties bodyProperties2 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties() { FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill11 = new A.SolidFill();

            A.SchemeColor schemeColor19 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation11 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor19.Append(luminanceModulation11);
            schemeColor19.Append(luminanceOffset3);

            solidFill11.Append(schemeColor19);
            A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont4 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties2.Append(solidFill11);
            defaultRunProperties2.Append(latinFont4);
            defaultRunProperties2.Append(eastAsianFont4);
            defaultRunProperties2.Append(complexScriptFont4);

            paragraphProperties2.Append(defaultRunProperties2);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(endParagraphRunProperties1);

            textProperties2.Append(bodyProperties2);
            textProperties2.Append(listStyle2);
            textProperties2.Append(paragraph2);

            title1.Append(chartText1);
            title1.Append(layout2);
            title1.Append(overlay1);
            title1.Append(chartShapeProperties6);
            title1.Append(textProperties2);
            C.NumberingFormat numberingFormat1 = new C.NumberingFormat() { FormatCode = "#,##0.000", SourceLinked = true };
            C.MajorTickMark majorTickMark1 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
            C.MinorTickMark minorTickMark1 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
            C.TickLabelPosition tickLabelPosition1 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.Low };
            C.CrossingAxis crossingAxis1 = new C.CrossingAxis() { Val = (UInt32Value)375001960U };
            C.Crosses crosses1 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
            C.CrossBetween crossBetween1 = new C.CrossBetween() { Val = C.CrossBetweenValues.MidpointCategory };
            C.MajorUnit majorUnit1 = new C.MajorUnit() { Val = 1D };

            valueAxis1.Append(axisId3);
            valueAxis1.Append(scaling1);
            //valueAxis1.Append(delete1);
            valueAxis1.Append(axisPosition1);
            valueAxis1.Append(majorGridlines1);
            valueAxis1.Append(title1);
            //valueAxis1.Append(numberingFormat1);
            valueAxis1.Append(majorTickMark1);
            valueAxis1.Append(minorTickMark1);
            valueAxis1.Append(tickLabelPosition1);
            valueAxis1.Append(crossingAxis1);
            valueAxis1.Append(crosses1);
            valueAxis1.Append(crossBetween1);
            valueAxis1.Append(majorUnit1);

            C.ValueAxis valueAxis2 = new C.ValueAxis();
            C.AxisId axisId4 = new C.AxisId() { Val = (UInt32Value)375001960U };

            C.Scaling scaling2 = new C.Scaling();
            C.Orientation orientation2 = new C.Orientation() { Val = C.OrientationValues.MinMax };
            C.MaxAxisValue maxAxisValue2 = new C.MaxAxisValue() { Val = 1D };

            scaling2.Append(orientation2);
            scaling2.Append(maxAxisValue2);
            C.Delete delete2 = new C.Delete() { Val = true };
            C.AxisPosition axisPosition2 = new C.AxisPosition() { Val = C.AxisPositionValues.Left };

            C.MajorGridlines majorGridlines2 = new C.MajorGridlines();

            C.ChartShapeProperties chartShapeProperties7 = new C.ChartShapeProperties();

            A.Outline outline10 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill12 = new A.SolidFill();

            A.SchemeColor schemeColor20 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation12 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset4 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor20.Append(luminanceModulation12);
            schemeColor20.Append(luminanceOffset4);

            solidFill12.Append(schemeColor20);
            A.Round round4 = new A.Round();

            outline10.Append(solidFill12);
            outline10.Append(round4);
            A.EffectList effectList10 = new A.EffectList();

            chartShapeProperties7.Append(outline10);
            chartShapeProperties7.Append(effectList10);

            majorGridlines2.Append(chartShapeProperties7);

            C.Title title2 = new C.Title();

            C.ChartText chartText2 = new C.ChartText();

            C.RichText richText2 = new C.RichText();
            A.BodyProperties bodyProperties3 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties() { FontSize = 1000, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill13 = new A.SolidFill();

            A.SchemeColor schemeColor21 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation13 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset5 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor21.Append(luminanceModulation13);
            schemeColor21.Append(luminanceOffset5);

            solidFill13.Append(schemeColor21);
            A.LatinFont latinFont5 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont5 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont5 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties3.Append(solidFill13);
            defaultRunProperties3.Append(latinFont5);
            defaultRunProperties3.Append(eastAsianFont5);
            defaultRunProperties3.Append(complexScriptFont5);

            paragraphProperties3.Append(defaultRunProperties3);

            A.Run run2 = new A.Run();
            A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-US", Bold = true };
            A.Text text2 = new A.Text();
            text2.Text = "Sales";

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run2);

            richText2.Append(bodyProperties3);
            richText2.Append(listStyle3);
            richText2.Append(paragraph3);

            chartText2.Append(richText2);

            C.Layout layout3 = new C.Layout();

            C.ManualLayout manualLayout3 = new C.ManualLayout();
            C.LeftMode leftMode3 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
            C.TopMode topMode3 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
            C.Left left3 = new C.Left() { Val = 4.5307817535466303E-2D };
            C.Top top3 = new C.Top() { Val = 0.48469203544678863D };

            manualLayout3.Append(leftMode3);
            manualLayout3.Append(topMode3);
            manualLayout3.Append(left3);
            manualLayout3.Append(top3);

            layout3.Append(manualLayout3);
            C.Overlay overlay2 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties8 = new C.ChartShapeProperties();
            A.NoFill noFill7 = new A.NoFill();

            A.Outline outline11 = new A.Outline();
            A.NoFill noFill8 = new A.NoFill();

            outline11.Append(noFill8);
            A.EffectList effectList11 = new A.EffectList();

            chartShapeProperties8.Append(noFill7);
            chartShapeProperties8.Append(outline11);
            chartShapeProperties8.Append(effectList11);

            C.TextProperties textProperties3 = new C.TextProperties();
            A.BodyProperties bodyProperties4 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties() { FontSize = 1000, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill14 = new A.SolidFill();

            A.SchemeColor schemeColor22 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation14 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset6 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor22.Append(luminanceModulation14);
            schemeColor22.Append(luminanceOffset6);

            solidFill14.Append(schemeColor22);
            A.LatinFont latinFont6 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont6 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties4.Append(solidFill14);
            defaultRunProperties4.Append(latinFont6);
            defaultRunProperties4.Append(eastAsianFont6);
            defaultRunProperties4.Append(complexScriptFont6);

            paragraphProperties4.Append(defaultRunProperties4);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(endParagraphRunProperties2);

            textProperties3.Append(bodyProperties4);
            textProperties3.Append(listStyle4);
            textProperties3.Append(paragraph4);

            title2.Append(chartText2);
            title2.Append(layout3);
            title2.Append(overlay2);
            title2.Append(chartShapeProperties8);
            title2.Append(textProperties3);
            C.NumberingFormat numberingFormat2 = new C.NumberingFormat() { FormatCode = "#,##0.000", SourceLinked = true };
            C.MajorTickMark majorTickMark2 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
            C.MinorTickMark minorTickMark2 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
            C.TickLabelPosition tickLabelPosition2 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };
            C.CrossingAxis crossingAxis2 = new C.CrossingAxis() { Val = (UInt32Value)375001568U };
            C.Crosses crosses2 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
            C.CrossBetween crossBetween2 = new C.CrossBetween() { Val = C.CrossBetweenValues.MidpointCategory };
            C.MajorUnit majorUnit2 = new C.MajorUnit() { Val = 1D };

            valueAxis2.Append(axisId4);
            valueAxis2.Append(scaling2);
            valueAxis2.Append(delete2);
            valueAxis2.Append(axisPosition2);
            valueAxis2.Append(majorGridlines2);
            valueAxis2.Append(title2);
            valueAxis2.Append(numberingFormat2);
            valueAxis2.Append(majorTickMark2);
            valueAxis2.Append(minorTickMark2);
            valueAxis2.Append(tickLabelPosition2);
            valueAxis2.Append(crossingAxis2);
            valueAxis2.Append(crosses2);
            valueAxis2.Append(crossBetween2);
            valueAxis2.Append(majorUnit2);

            C.ShapeProperties shapeProperties1 = new C.ShapeProperties();
            A.NoFill noFill9 = new A.NoFill();

            A.Outline outline12 = new A.Outline();

            A.SolidFill solidFill15 = new A.SolidFill();
            A.SchemeColor schemeColor23 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill15.Append(schemeColor23);

            outline12.Append(solidFill15);
            A.EffectList effectList12 = new A.EffectList();

            shapeProperties1.Append(noFill9);
            shapeProperties1.Append(outline12);
            shapeProperties1.Append(effectList12);

            plotArea1.Append(layout1);
            plotArea1.Append(scatterChart1);
            plotArea1.Append(valueAxis1);
            plotArea1.Append(valueAxis2);
            plotArea1.Append(shapeProperties1);
            C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };
            C.DisplayBlanksAs displayBlanksAs1 = new C.DisplayBlanksAs() { Val = C.DisplayBlanksAsValues.Gap };
            C.ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new C.ShowDataLabelsOverMaximum() { Val = false };

            //chart1.Append(autoTitleDeleted1);
            chart1.Append(plotArea1);
            //chart1.Append(plotVisibleOnly1);
            //chart1.Append(displayBlanksAs1);
            //chart1.Append(showDataLabelsOverMaximum1);

            C.ShapeProperties shapeProperties2 = new C.ShapeProperties();

            A.SolidFill solidFill16 = new A.SolidFill();
            A.SchemeColor schemeColor24 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill16.Append(schemeColor24);

            A.Outline outline13 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill17 = new A.SolidFill();
            A.SchemeColor schemeColor25 = new A.SchemeColor() { Val = A.SchemeColorValues.Background2 };

            solidFill17.Append(schemeColor25);
            A.Round round5 = new A.Round();

            outline13.Append(solidFill17);
            outline13.Append(round5);
            A.EffectList effectList13 = new A.EffectList();

            shapeProperties2.Append(solidFill16);
            shapeProperties2.Append(outline13);
            shapeProperties2.Append(effectList13);

            C.TextProperties textProperties4 = new C.TextProperties();
            A.BodyProperties bodyProperties5 = new A.BodyProperties();
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties();

            paragraphProperties5.Append(defaultRunProperties5);
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(endParagraphRunProperties3);

            textProperties4.Append(bodyProperties5);
            textProperties4.Append(listStyle5);
            textProperties4.Append(paragraph5);

            C.PrintSettings printSettings1 = new C.PrintSettings();
            C.HeaderFooter headerFooter1 = new C.HeaderFooter();
            C.PageMargins pageMargins2 = new C.PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            C.PageSetup pageSetup2 = new C.PageSetup();

            printSettings1.Append(headerFooter1);
            printSettings1.Append(pageMargins2);
            printSettings1.Append(pageSetup2);

            //chartSpace1.Append(date19041);
            //chartSpace1.Append(editingLanguage1);
            chartSpace1.Append(roundedCorners1);
            //chartSpace1.Append(alternateContent2);
            chartSpace1.Append(chart1);
            //chartSpace1.Append(shapeProperties2);
            //chartSpace1.Append(textProperties4);
            //chartSpace1.Append(printSettings1);

            chartPart1.ChartSpace = chartSpace1;
        }

        // Generates content of sharedStringTablePart1.
        private void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)4U, UniqueCount = (UInt32Value)4U };

            SharedStringItem sharedStringItem1 = new SharedStringItem();
            Text text3 = new Text();
            text3.Text = "Obs #";

            sharedStringItem1.Append(text3);

            SharedStringItem sharedStringItem2 = new SharedStringItem();
            Text text4 = new Text();
            text4.Text = "SALES";

            sharedStringItem2.Append(text4);

            SharedStringItem sharedStringItem3 = new SharedStringItem();
            Text text5 = new Text();
            text5.Text = "Base Observations";

            sharedStringItem3.Append(text5);

            SharedStringItem sharedStringItem4 = new SharedStringItem();
            Text text6 = new Text();
            text6.Text = "Projected Observations";

            sharedStringItem4.Append(text6);

            sharedStringTable1.Append(sharedStringItem1);
            sharedStringTable1.Append(sharedStringItem2);
            sharedStringTable1.Append(sharedStringItem3);
            sharedStringTable1.Append(sharedStringItem4);

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }

        // Generates content of workbookStylesPart1.
        private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            NumberingFormats numberingFormats1 = new NumberingFormats() { Count = (UInt32Value)1U };
            NumberingFormat numberingFormat3 = new NumberingFormat() { NumberFormatId = (UInt32Value)164U, FormatCode = "#,##0.000" };

            numberingFormats1.Append(numberingFormat3);

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)3U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme2);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme3);

            Font font3 = new Font();
            Bold bold2 = new Bold();
            FontSize fontSize3 = new FontSize() { Val = 14D };
            Color color3 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName3 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme4 = new FontScheme() { Val = FontSchemeValues.Minor };

            font3.Append(bold2);
            font3.Append(fontSize3);
            font3.Append(color3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);
            font3.Append(fontScheme4);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);

            Fills fills1 = new Fills() { Count = (UInt32Value)4U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            Fill fill3 = new Fill();

            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Theme = (UInt32Value)8U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);

            fill3.Append(patternFill3);

            Fill fill4 = new Fill();

            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Theme = (UInt32Value)2U, Tint = -9.9978637043366805E-2D };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill4.Append(foregroundColor2);
            patternFill4.Append(backgroundColor2);

            fill4.Append(patternFill4);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);
            fills1.Append(fill4);

            Borders borders1 = new Borders() { Count = (UInt32Value)1U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            borders1.Append(border1);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)6U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)164U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true };
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)164U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true };

            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment1 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, TextRotation = (UInt32Value)180U };

            cellFormat6.Append(alignment1);

            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment2 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, TextRotation = (UInt32Value)180U };

            cellFormat7.Append(alignment2);

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);
            cellFormats1.Append(cellFormat6);
            cellFormats1.Append(cellFormat7);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(numberingFormats1);
            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }

    }
}
