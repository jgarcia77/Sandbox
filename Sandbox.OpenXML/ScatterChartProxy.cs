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
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart1);

            ConnectionsPart connectionsPart1 = workbookPart1.AddNewPart<ConnectionsPart>("rId3");
            GenerateConnectionsPart1Content(connectionsPart1);

            ThemePart themePart1 = workbookPart1.AddNewPart<ThemePart>("rId2");
            GenerateThemePart1Content(themePart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPart1Content(worksheetPart1);

            QueryTablePart queryTablePart1 = worksheetPart1.AddNewPart<QueryTablePart>("rId3");
            GenerateQueryTablePart1Content(queryTablePart1);

            DrawingsPart drawingsPart1 = worksheetPart1.AddNewPart<DrawingsPart>("rId2");
            GenerateDrawingsPart1Content(drawingsPart1);

            ChartPart chartPart1 = drawingsPart1.AddNewPart<ChartPart>("rId1");
            GenerateChartPart1Content(chartPart1);

            ChartColorStylePart chartColorStylePart1 = chartPart1.AddNewPart<ChartColorStylePart>("rId2");
            GenerateChartColorStylePart1Content(chartColorStylePart1);

            ChartStylePart chartStylePart1 = chartPart1.AddNewPart<ChartStylePart>("rId1");
            GenerateChartStylePart1Content(chartStylePart1);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1 = worksheetPart1.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart1Content(spreadsheetPrinterSettingsPart1);

            SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId5");
            GenerateSharedStringTablePart1Content(sharedStringTablePart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId4");
            GenerateWorkbookStylesPart1Content(workbookStylesPart1);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Excel";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)4U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Worksheets";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            Vt.Variant variant3 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Named Ranges";

            variant3.Append(vTLPSTR2);

            Vt.Variant variant4 = new Vt.Variant();
            Vt.VTInt32 vTInt322 = new Vt.VTInt32();
            vTInt322.Text = "1";

            variant4.Append(vTInt322);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);
            vTVector1.Append(variant3);
            vTVector1.Append(variant4);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)2U };
            Vt.VTLPSTR vTLPSTR3 = new Vt.VTLPSTR();
            vTLPSTR3.Text = "Scatter Matrix";
            Vt.VTLPSTR vTLPSTR4 = new Vt.VTLPSTR();
            vTLPSTR4.Text = "\'Scatter Matrix\'!Multivariate_Data";

            vTVector2.Append(vTLPSTR3);
            vTVector2.Append(vTLPSTR4);

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = "Magenic Technologies Inc.";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "15.0300";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
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

        // Generates content of connectionsPart1.
        private void GenerateConnectionsPart1Content(ConnectionsPart connectionsPart1)
        {
            Connections connections1 = new Connections();

            Connection connection1 = new Connection() { Id = (UInt32Value)1U, Name = "Multivariate Data", Type = (UInt32Value)6U, RefreshedVersion = 5, Background = true, SaveData = true };

            TextProperties textProperties1 = new TextProperties() { CodePage = (UInt32Value)437U, SourceFile = "C:\\Users\\josueg\\Documents\\Projects\\Deloitte\\STAR Modernization - Magenic Documents\\Multivariate Data.txt" };

            TextFields textFields1 = new TextFields() { Count = (UInt32Value)13U };
            TextField textField1 = new TextField();
            TextField textField2 = new TextField();
            TextField textField3 = new TextField();
            TextField textField4 = new TextField();
            TextField textField5 = new TextField();
            TextField textField6 = new TextField();
            TextField textField7 = new TextField();
            TextField textField8 = new TextField();
            TextField textField9 = new TextField();
            TextField textField10 = new TextField();
            TextField textField11 = new TextField();
            TextField textField12 = new TextField();
            TextField textField13 = new TextField();

            textFields1.Append(textField1);
            textFields1.Append(textField2);
            textFields1.Append(textField3);
            textFields1.Append(textField4);
            textFields1.Append(textField5);
            textFields1.Append(textField6);
            textFields1.Append(textField7);
            textFields1.Append(textField8);
            textFields1.Append(textField9);
            textFields1.Append(textField10);
            textFields1.Append(textField11);
            textFields1.Append(textField12);
            textFields1.Append(textField13);

            textProperties1.Append(textFields1);

            connection1.Append(textProperties1);

            connections1.Append(connection1);

            connectionsPart1.Connections = connections1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint1 = new A.Tint() { Val = 67000 };

            schemeColor2.Append(luminanceModulation1);
            schemeColor2.Append(saturationModulation1);
            schemeColor2.Append(tint1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor3.Append(luminanceModulation2);
            schemeColor3.Append(saturationModulation2);
            schemeColor3.Append(tint2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(saturationModulation3);
            schemeColor4.Append(tint3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor5.Append(saturationModulation4);
            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(tint4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor6.Append(saturationModulation5);
            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(shade1);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor7.Append(luminanceModulation6);
            schemeColor7.Append(saturationModulation6);
            schemeColor7.Append(shade2);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);
            outline3.Append(miter3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill6.Append(schemeColor12);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor13.Append(tint6);
            schemeColor13.Append(saturationModulation8);
            schemeColor13.Append(shade3);
            schemeColor13.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor13);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor14.Append(tint7);
            schemeColor14.Append(saturationModulation9);
            schemeColor14.Append(shade4);
            schemeColor14.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor14);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor15.Append(shade5);
            schemeColor15.Append(saturationModulation10);

            gradientStop9.Append(schemeColor15);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(solidFill6);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
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

        // Generates content of queryTablePart1.
        private void GenerateQueryTablePart1Content(QueryTablePart queryTablePart1)
        {
            QueryTable queryTable1 = new QueryTable() { Name = "Multivariate Data", ConnectionId = (UInt32Value)1U, AutoFormatId = (UInt32Value)16U, ApplyNumberFormats = false, ApplyBorderFormats = false, ApplyFontFormats = false, ApplyPatternFormats = false, ApplyAlignmentFormats = false, ApplyWidthHeightFormats = false };

            queryTablePart1.QueryTable = queryTable1;
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

        // Generates content of chartColorStylePart1.
        private void GenerateChartColorStylePart1Content(ChartColorStylePart chartColorStylePart1)
        {
            Cs.ColorStyle colorStyle1 = new Cs.ColorStyle() { Method = "cycle", Id = (UInt32Value)10U };
            colorStyle1.AddNamespaceDeclaration("cs", "http://schemas.microsoft.com/office/drawing/2012/chartStyle");
            colorStyle1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.SchemeColor schemeColor26 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.SchemeColor schemeColor27 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent2 };
            A.SchemeColor schemeColor28 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent3 };
            A.SchemeColor schemeColor29 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent4 };
            A.SchemeColor schemeColor30 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent5 };
            A.SchemeColor schemeColor31 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent6 };
            Cs.ColorStyleVariation colorStyleVariation1 = new Cs.ColorStyleVariation();

            Cs.ColorStyleVariation colorStyleVariation2 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation15 = new A.LuminanceModulation() { Val = 60000 };

            colorStyleVariation2.Append(luminanceModulation15);

            Cs.ColorStyleVariation colorStyleVariation3 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation16 = new A.LuminanceModulation() { Val = 80000 };
            A.LuminanceOffset luminanceOffset7 = new A.LuminanceOffset() { Val = 20000 };

            colorStyleVariation3.Append(luminanceModulation16);
            colorStyleVariation3.Append(luminanceOffset7);

            Cs.ColorStyleVariation colorStyleVariation4 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation17 = new A.LuminanceModulation() { Val = 80000 };

            colorStyleVariation4.Append(luminanceModulation17);

            Cs.ColorStyleVariation colorStyleVariation5 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation18 = new A.LuminanceModulation() { Val = 60000 };
            A.LuminanceOffset luminanceOffset8 = new A.LuminanceOffset() { Val = 40000 };

            colorStyleVariation5.Append(luminanceModulation18);
            colorStyleVariation5.Append(luminanceOffset8);

            Cs.ColorStyleVariation colorStyleVariation6 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation19 = new A.LuminanceModulation() { Val = 50000 };

            colorStyleVariation6.Append(luminanceModulation19);

            Cs.ColorStyleVariation colorStyleVariation7 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation20 = new A.LuminanceModulation() { Val = 70000 };
            A.LuminanceOffset luminanceOffset9 = new A.LuminanceOffset() { Val = 30000 };

            colorStyleVariation7.Append(luminanceModulation20);
            colorStyleVariation7.Append(luminanceOffset9);

            Cs.ColorStyleVariation colorStyleVariation8 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation21 = new A.LuminanceModulation() { Val = 70000 };

            colorStyleVariation8.Append(luminanceModulation21);

            Cs.ColorStyleVariation colorStyleVariation9 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation22 = new A.LuminanceModulation() { Val = 50000 };
            A.LuminanceOffset luminanceOffset10 = new A.LuminanceOffset() { Val = 50000 };

            colorStyleVariation9.Append(luminanceModulation22);
            colorStyleVariation9.Append(luminanceOffset10);

            colorStyle1.Append(schemeColor26);
            colorStyle1.Append(schemeColor27);
            colorStyle1.Append(schemeColor28);
            colorStyle1.Append(schemeColor29);
            colorStyle1.Append(schemeColor30);
            colorStyle1.Append(schemeColor31);
            colorStyle1.Append(colorStyleVariation1);
            colorStyle1.Append(colorStyleVariation2);
            colorStyle1.Append(colorStyleVariation3);
            colorStyle1.Append(colorStyleVariation4);
            colorStyle1.Append(colorStyleVariation5);
            colorStyle1.Append(colorStyleVariation6);
            colorStyle1.Append(colorStyleVariation7);
            colorStyle1.Append(colorStyleVariation8);
            colorStyle1.Append(colorStyleVariation9);

            chartColorStylePart1.ColorStyle = colorStyle1;
        }

        // Generates content of chartStylePart1.
        private void GenerateChartStylePart1Content(ChartStylePart chartStylePart1)
        {
            Cs.ChartStyle chartStyle1 = new Cs.ChartStyle() { Id = (UInt32Value)240U };
            chartStyle1.AddNamespaceDeclaration("cs", "http://schemas.microsoft.com/office/drawing/2012/chartStyle");
            chartStyle1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Cs.AxisTitle axisTitle1 = new Cs.AxisTitle();
            Cs.LineReference lineReference1 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference1 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference1 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference1 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor32 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation23 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset11 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor32.Append(luminanceModulation23);
            schemeColor32.Append(luminanceOffset11);

            fontReference1.Append(schemeColor32);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType1 = new Cs.TextCharacterPropertiesType() { FontSize = 1000, Kerning = 1200 };

            axisTitle1.Append(lineReference1);
            axisTitle1.Append(fillReference1);
            axisTitle1.Append(effectReference1);
            axisTitle1.Append(fontReference1);
            axisTitle1.Append(textCharacterPropertiesType1);

            Cs.CategoryAxis categoryAxis1 = new Cs.CategoryAxis();
            Cs.LineReference lineReference2 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference2 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference2 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference2 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor33 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation24 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset12 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor33.Append(luminanceModulation24);
            schemeColor33.Append(luminanceOffset12);

            fontReference2.Append(schemeColor33);

            Cs.ShapeProperties shapeProperties3 = new Cs.ShapeProperties();

            A.Outline outline14 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill18 = new A.SolidFill();

            A.SchemeColor schemeColor34 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation25 = new A.LuminanceModulation() { Val = 25000 };
            A.LuminanceOffset luminanceOffset13 = new A.LuminanceOffset() { Val = 75000 };

            schemeColor34.Append(luminanceModulation25);
            schemeColor34.Append(luminanceOffset13);

            solidFill18.Append(schemeColor34);
            A.Round round6 = new A.Round();

            outline14.Append(solidFill18);
            outline14.Append(round6);

            shapeProperties3.Append(outline14);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType2 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            categoryAxis1.Append(lineReference2);
            categoryAxis1.Append(fillReference2);
            categoryAxis1.Append(effectReference2);
            categoryAxis1.Append(fontReference2);
            categoryAxis1.Append(shapeProperties3);
            categoryAxis1.Append(textCharacterPropertiesType2);

            Cs.ChartArea chartArea1 = new Cs.ChartArea() { Modifiers = new ListValue<StringValue>() { InnerText = "allowNoFillOverride allowNoLineOverride" } };
            Cs.LineReference lineReference3 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference3 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference3 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference3 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor35 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference3.Append(schemeColor35);

            Cs.ShapeProperties shapeProperties4 = new Cs.ShapeProperties();

            A.SolidFill solidFill19 = new A.SolidFill();
            A.SchemeColor schemeColor36 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill19.Append(schemeColor36);

            A.Outline outline15 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill20 = new A.SolidFill();

            A.SchemeColor schemeColor37 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation26 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset14 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor37.Append(luminanceModulation26);
            schemeColor37.Append(luminanceOffset14);

            solidFill20.Append(schemeColor37);
            A.Round round7 = new A.Round();

            outline15.Append(solidFill20);
            outline15.Append(round7);

            shapeProperties4.Append(solidFill19);
            shapeProperties4.Append(outline15);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType3 = new Cs.TextCharacterPropertiesType() { FontSize = 1000, Kerning = 1200 };

            chartArea1.Append(lineReference3);
            chartArea1.Append(fillReference3);
            chartArea1.Append(effectReference3);
            chartArea1.Append(fontReference3);
            chartArea1.Append(shapeProperties4);
            chartArea1.Append(textCharacterPropertiesType3);

            Cs.DataLabel dataLabel1 = new Cs.DataLabel();
            Cs.LineReference lineReference4 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference4 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference4 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference4 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor38 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation27 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset15 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor38.Append(luminanceModulation27);
            schemeColor38.Append(luminanceOffset15);

            fontReference4.Append(schemeColor38);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType4 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            dataLabel1.Append(lineReference4);
            dataLabel1.Append(fillReference4);
            dataLabel1.Append(effectReference4);
            dataLabel1.Append(fontReference4);
            dataLabel1.Append(textCharacterPropertiesType4);

            Cs.DataLabelCallout dataLabelCallout1 = new Cs.DataLabelCallout();
            Cs.LineReference lineReference5 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference5 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference5 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference5 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor39 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation28 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset16 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor39.Append(luminanceModulation28);
            schemeColor39.Append(luminanceOffset16);

            fontReference5.Append(schemeColor39);

            Cs.ShapeProperties shapeProperties5 = new Cs.ShapeProperties();

            A.SolidFill solidFill21 = new A.SolidFill();
            A.SchemeColor schemeColor40 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            solidFill21.Append(schemeColor40);

            A.Outline outline16 = new A.Outline();

            A.SolidFill solidFill22 = new A.SolidFill();

            A.SchemeColor schemeColor41 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation29 = new A.LuminanceModulation() { Val = 25000 };
            A.LuminanceOffset luminanceOffset17 = new A.LuminanceOffset() { Val = 75000 };

            schemeColor41.Append(luminanceModulation29);
            schemeColor41.Append(luminanceOffset17);

            solidFill22.Append(schemeColor41);

            outline16.Append(solidFill22);

            shapeProperties5.Append(solidFill21);
            shapeProperties5.Append(outline16);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType5 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            Cs.TextBodyProperties textBodyProperties1 = new Cs.TextBodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Clip, HorizontalOverflow = A.TextHorizontalOverflowValues.Clip, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 36576, TopInset = 18288, RightInset = 36576, BottomInset = 18288, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

            textBodyProperties1.Append(shapeAutoFit1);

            dataLabelCallout1.Append(lineReference5);
            dataLabelCallout1.Append(fillReference5);
            dataLabelCallout1.Append(effectReference5);
            dataLabelCallout1.Append(fontReference5);
            dataLabelCallout1.Append(shapeProperties5);
            dataLabelCallout1.Append(textCharacterPropertiesType5);
            dataLabelCallout1.Append(textBodyProperties1);

            Cs.DataPoint dataPoint1 = new Cs.DataPoint();
            Cs.LineReference lineReference6 = new Cs.LineReference() { Index = (UInt32Value)0U };

            Cs.FillReference fillReference6 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.StyleColor styleColor1 = new Cs.StyleColor() { Val = "auto" };

            fillReference6.Append(styleColor1);
            Cs.EffectReference effectReference6 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference6 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor42 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference6.Append(schemeColor42);

            dataPoint1.Append(lineReference6);
            dataPoint1.Append(fillReference6);
            dataPoint1.Append(effectReference6);
            dataPoint1.Append(fontReference6);

            Cs.DataPoint3D dataPoint3D1 = new Cs.DataPoint3D();
            Cs.LineReference lineReference7 = new Cs.LineReference() { Index = (UInt32Value)0U };

            Cs.FillReference fillReference7 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.StyleColor styleColor2 = new Cs.StyleColor() { Val = "auto" };

            fillReference7.Append(styleColor2);
            Cs.EffectReference effectReference7 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference7 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor43 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference7.Append(schemeColor43);

            dataPoint3D1.Append(lineReference7);
            dataPoint3D1.Append(fillReference7);
            dataPoint3D1.Append(effectReference7);
            dataPoint3D1.Append(fontReference7);

            Cs.DataPointLine dataPointLine1 = new Cs.DataPointLine();

            Cs.LineReference lineReference8 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor3 = new Cs.StyleColor() { Val = "auto" };

            lineReference8.Append(styleColor3);
            Cs.FillReference fillReference8 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.EffectReference effectReference8 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference8 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor44 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference8.Append(schemeColor44);

            Cs.ShapeProperties shapeProperties6 = new Cs.ShapeProperties();

            A.Outline outline17 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill23 = new A.SolidFill();
            A.SchemeColor schemeColor45 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill23.Append(schemeColor45);
            A.Round round8 = new A.Round();

            outline17.Append(solidFill23);
            outline17.Append(round8);

            shapeProperties6.Append(outline17);

            dataPointLine1.Append(lineReference8);
            dataPointLine1.Append(fillReference8);
            dataPointLine1.Append(effectReference8);
            dataPointLine1.Append(fontReference8);
            dataPointLine1.Append(shapeProperties6);

            Cs.DataPointMarker dataPointMarker1 = new Cs.DataPointMarker();

            Cs.LineReference lineReference9 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor4 = new Cs.StyleColor() { Val = "auto" };

            lineReference9.Append(styleColor4);

            Cs.FillReference fillReference9 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.StyleColor styleColor5 = new Cs.StyleColor() { Val = "auto" };

            fillReference9.Append(styleColor5);
            Cs.EffectReference effectReference9 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference9 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor46 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference9.Append(schemeColor46);

            Cs.ShapeProperties shapeProperties7 = new Cs.ShapeProperties();

            A.Outline outline18 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill24 = new A.SolidFill();
            A.SchemeColor schemeColor47 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill24.Append(schemeColor47);

            outline18.Append(solidFill24);

            shapeProperties7.Append(outline18);

            dataPointMarker1.Append(lineReference9);
            dataPointMarker1.Append(fillReference9);
            dataPointMarker1.Append(effectReference9);
            dataPointMarker1.Append(fontReference9);
            dataPointMarker1.Append(shapeProperties7);
            Cs.MarkerLayoutProperties markerLayoutProperties1 = new Cs.MarkerLayoutProperties() { Symbol = Cs.MarkerStyle.Circle, Size = 5 };

            Cs.DataPointWireframe dataPointWireframe1 = new Cs.DataPointWireframe();

            Cs.LineReference lineReference10 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor6 = new Cs.StyleColor() { Val = "auto" };

            lineReference10.Append(styleColor6);
            Cs.FillReference fillReference10 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference10 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference10 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor48 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference10.Append(schemeColor48);

            Cs.ShapeProperties shapeProperties8 = new Cs.ShapeProperties();

            A.Outline outline19 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill25 = new A.SolidFill();
            A.SchemeColor schemeColor49 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill25.Append(schemeColor49);
            A.Round round9 = new A.Round();

            outline19.Append(solidFill25);
            outline19.Append(round9);

            shapeProperties8.Append(outline19);

            dataPointWireframe1.Append(lineReference10);
            dataPointWireframe1.Append(fillReference10);
            dataPointWireframe1.Append(effectReference10);
            dataPointWireframe1.Append(fontReference10);
            dataPointWireframe1.Append(shapeProperties8);

            Cs.DataTableStyle dataTableStyle1 = new Cs.DataTableStyle();
            Cs.LineReference lineReference11 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference11 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference11 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference11 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor50 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation30 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset18 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor50.Append(luminanceModulation30);
            schemeColor50.Append(luminanceOffset18);

            fontReference11.Append(schemeColor50);

            Cs.ShapeProperties shapeProperties9 = new Cs.ShapeProperties();
            A.NoFill noFill10 = new A.NoFill();

            A.Outline outline20 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill26 = new A.SolidFill();

            A.SchemeColor schemeColor51 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation31 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset19 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor51.Append(luminanceModulation31);
            schemeColor51.Append(luminanceOffset19);

            solidFill26.Append(schemeColor51);
            A.Round round10 = new A.Round();

            outline20.Append(solidFill26);
            outline20.Append(round10);

            shapeProperties9.Append(noFill10);
            shapeProperties9.Append(outline20);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType6 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            dataTableStyle1.Append(lineReference11);
            dataTableStyle1.Append(fillReference11);
            dataTableStyle1.Append(effectReference11);
            dataTableStyle1.Append(fontReference11);
            dataTableStyle1.Append(shapeProperties9);
            dataTableStyle1.Append(textCharacterPropertiesType6);

            Cs.DownBar downBar1 = new Cs.DownBar();
            Cs.LineReference lineReference12 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference12 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference12 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference12 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor52 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference12.Append(schemeColor52);

            Cs.ShapeProperties shapeProperties10 = new Cs.ShapeProperties();

            A.SolidFill solidFill27 = new A.SolidFill();

            A.SchemeColor schemeColor53 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation32 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset20 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor53.Append(luminanceModulation32);
            schemeColor53.Append(luminanceOffset20);

            solidFill27.Append(schemeColor53);

            A.Outline outline21 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill28 = new A.SolidFill();

            A.SchemeColor schemeColor54 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation33 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset21 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor54.Append(luminanceModulation33);
            schemeColor54.Append(luminanceOffset21);

            solidFill28.Append(schemeColor54);
            A.Round round11 = new A.Round();

            outline21.Append(solidFill28);
            outline21.Append(round11);

            shapeProperties10.Append(solidFill27);
            shapeProperties10.Append(outline21);

            downBar1.Append(lineReference12);
            downBar1.Append(fillReference12);
            downBar1.Append(effectReference12);
            downBar1.Append(fontReference12);
            downBar1.Append(shapeProperties10);

            Cs.DropLine dropLine1 = new Cs.DropLine();
            Cs.LineReference lineReference13 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference13 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference13 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference13 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor55 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference13.Append(schemeColor55);

            Cs.ShapeProperties shapeProperties11 = new Cs.ShapeProperties();

            A.Outline outline22 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill29 = new A.SolidFill();

            A.SchemeColor schemeColor56 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation34 = new A.LuminanceModulation() { Val = 35000 };
            A.LuminanceOffset luminanceOffset22 = new A.LuminanceOffset() { Val = 65000 };

            schemeColor56.Append(luminanceModulation34);
            schemeColor56.Append(luminanceOffset22);

            solidFill29.Append(schemeColor56);
            A.Round round12 = new A.Round();

            outline22.Append(solidFill29);
            outline22.Append(round12);

            shapeProperties11.Append(outline22);

            dropLine1.Append(lineReference13);
            dropLine1.Append(fillReference13);
            dropLine1.Append(effectReference13);
            dropLine1.Append(fontReference13);
            dropLine1.Append(shapeProperties11);

            Cs.ErrorBar errorBar1 = new Cs.ErrorBar();
            Cs.LineReference lineReference14 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference14 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference14 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference14 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor57 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference14.Append(schemeColor57);

            Cs.ShapeProperties shapeProperties12 = new Cs.ShapeProperties();

            A.Outline outline23 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill30 = new A.SolidFill();

            A.SchemeColor schemeColor58 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation35 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset23 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor58.Append(luminanceModulation35);
            schemeColor58.Append(luminanceOffset23);

            solidFill30.Append(schemeColor58);
            A.Round round13 = new A.Round();

            outline23.Append(solidFill30);
            outline23.Append(round13);

            shapeProperties12.Append(outline23);

            errorBar1.Append(lineReference14);
            errorBar1.Append(fillReference14);
            errorBar1.Append(effectReference14);
            errorBar1.Append(fontReference14);
            errorBar1.Append(shapeProperties12);

            Cs.Floor floor1 = new Cs.Floor();
            Cs.LineReference lineReference15 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference15 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference15 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference15 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor59 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference15.Append(schemeColor59);

            Cs.ShapeProperties shapeProperties13 = new Cs.ShapeProperties();
            A.NoFill noFill11 = new A.NoFill();

            A.Outline outline24 = new A.Outline();
            A.NoFill noFill12 = new A.NoFill();

            outline24.Append(noFill12);

            shapeProperties13.Append(noFill11);
            shapeProperties13.Append(outline24);

            floor1.Append(lineReference15);
            floor1.Append(fillReference15);
            floor1.Append(effectReference15);
            floor1.Append(fontReference15);
            floor1.Append(shapeProperties13);

            Cs.GridlineMajor gridlineMajor1 = new Cs.GridlineMajor();
            Cs.LineReference lineReference16 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference16 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference16 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference16 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor60 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference16.Append(schemeColor60);

            Cs.ShapeProperties shapeProperties14 = new Cs.ShapeProperties();

            A.Outline outline25 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill31 = new A.SolidFill();

            A.SchemeColor schemeColor61 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation36 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset24 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor61.Append(luminanceModulation36);
            schemeColor61.Append(luminanceOffset24);

            solidFill31.Append(schemeColor61);
            A.Round round14 = new A.Round();

            outline25.Append(solidFill31);
            outline25.Append(round14);

            shapeProperties14.Append(outline25);

            gridlineMajor1.Append(lineReference16);
            gridlineMajor1.Append(fillReference16);
            gridlineMajor1.Append(effectReference16);
            gridlineMajor1.Append(fontReference16);
            gridlineMajor1.Append(shapeProperties14);

            Cs.GridlineMinor gridlineMinor1 = new Cs.GridlineMinor();
            Cs.LineReference lineReference17 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference17 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference17 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference17 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor62 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference17.Append(schemeColor62);

            Cs.ShapeProperties shapeProperties15 = new Cs.ShapeProperties();

            A.Outline outline26 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill32 = new A.SolidFill();

            A.SchemeColor schemeColor63 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation37 = new A.LuminanceModulation() { Val = 5000 };
            A.LuminanceOffset luminanceOffset25 = new A.LuminanceOffset() { Val = 95000 };

            schemeColor63.Append(luminanceModulation37);
            schemeColor63.Append(luminanceOffset25);

            solidFill32.Append(schemeColor63);
            A.Round round15 = new A.Round();

            outline26.Append(solidFill32);
            outline26.Append(round15);

            shapeProperties15.Append(outline26);

            gridlineMinor1.Append(lineReference17);
            gridlineMinor1.Append(fillReference17);
            gridlineMinor1.Append(effectReference17);
            gridlineMinor1.Append(fontReference17);
            gridlineMinor1.Append(shapeProperties15);

            Cs.HiLoLine hiLoLine1 = new Cs.HiLoLine();
            Cs.LineReference lineReference18 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference18 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference18 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference18 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor64 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference18.Append(schemeColor64);

            Cs.ShapeProperties shapeProperties16 = new Cs.ShapeProperties();

            A.Outline outline27 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill33 = new A.SolidFill();

            A.SchemeColor schemeColor65 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation38 = new A.LuminanceModulation() { Val = 50000 };
            A.LuminanceOffset luminanceOffset26 = new A.LuminanceOffset() { Val = 50000 };

            schemeColor65.Append(luminanceModulation38);
            schemeColor65.Append(luminanceOffset26);

            solidFill33.Append(schemeColor65);
            A.Round round16 = new A.Round();

            outline27.Append(solidFill33);
            outline27.Append(round16);

            shapeProperties16.Append(outline27);

            hiLoLine1.Append(lineReference18);
            hiLoLine1.Append(fillReference18);
            hiLoLine1.Append(effectReference18);
            hiLoLine1.Append(fontReference18);
            hiLoLine1.Append(shapeProperties16);

            Cs.LeaderLine leaderLine1 = new Cs.LeaderLine();
            Cs.LineReference lineReference19 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference19 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference19 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference19 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor66 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference19.Append(schemeColor66);

            Cs.ShapeProperties shapeProperties17 = new Cs.ShapeProperties();

            A.Outline outline28 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill34 = new A.SolidFill();

            A.SchemeColor schemeColor67 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation39 = new A.LuminanceModulation() { Val = 35000 };
            A.LuminanceOffset luminanceOffset27 = new A.LuminanceOffset() { Val = 65000 };

            schemeColor67.Append(luminanceModulation39);
            schemeColor67.Append(luminanceOffset27);

            solidFill34.Append(schemeColor67);
            A.Round round17 = new A.Round();

            outline28.Append(solidFill34);
            outline28.Append(round17);

            shapeProperties17.Append(outline28);

            leaderLine1.Append(lineReference19);
            leaderLine1.Append(fillReference19);
            leaderLine1.Append(effectReference19);
            leaderLine1.Append(fontReference19);
            leaderLine1.Append(shapeProperties17);

            Cs.LegendStyle legendStyle1 = new Cs.LegendStyle();
            Cs.LineReference lineReference20 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference20 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference20 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference20 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor68 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation40 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset28 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor68.Append(luminanceModulation40);
            schemeColor68.Append(luminanceOffset28);

            fontReference20.Append(schemeColor68);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType7 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            legendStyle1.Append(lineReference20);
            legendStyle1.Append(fillReference20);
            legendStyle1.Append(effectReference20);
            legendStyle1.Append(fontReference20);
            legendStyle1.Append(textCharacterPropertiesType7);

            Cs.PlotArea plotArea2 = new Cs.PlotArea() { Modifiers = new ListValue<StringValue>() { InnerText = "allowNoFillOverride allowNoLineOverride" } };
            Cs.LineReference lineReference21 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference21 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference21 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference21 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor69 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference21.Append(schemeColor69);

            plotArea2.Append(lineReference21);
            plotArea2.Append(fillReference21);
            plotArea2.Append(effectReference21);
            plotArea2.Append(fontReference21);

            Cs.PlotArea3D plotArea3D1 = new Cs.PlotArea3D() { Modifiers = new ListValue<StringValue>() { InnerText = "allowNoFillOverride allowNoLineOverride" } };
            Cs.LineReference lineReference22 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference22 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference22 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference22 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor70 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference22.Append(schemeColor70);

            plotArea3D1.Append(lineReference22);
            plotArea3D1.Append(fillReference22);
            plotArea3D1.Append(effectReference22);
            plotArea3D1.Append(fontReference22);

            Cs.SeriesAxis seriesAxis1 = new Cs.SeriesAxis();
            Cs.LineReference lineReference23 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference23 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference23 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference23 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor71 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation41 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset29 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor71.Append(luminanceModulation41);
            schemeColor71.Append(luminanceOffset29);

            fontReference23.Append(schemeColor71);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType8 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            seriesAxis1.Append(lineReference23);
            seriesAxis1.Append(fillReference23);
            seriesAxis1.Append(effectReference23);
            seriesAxis1.Append(fontReference23);
            seriesAxis1.Append(textCharacterPropertiesType8);

            Cs.SeriesLine seriesLine1 = new Cs.SeriesLine();
            Cs.LineReference lineReference24 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference24 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference24 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference24 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor72 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference24.Append(schemeColor72);

            Cs.ShapeProperties shapeProperties18 = new Cs.ShapeProperties();

            A.Outline outline29 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill35 = new A.SolidFill();

            A.SchemeColor schemeColor73 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation42 = new A.LuminanceModulation() { Val = 35000 };
            A.LuminanceOffset luminanceOffset30 = new A.LuminanceOffset() { Val = 65000 };

            schemeColor73.Append(luminanceModulation42);
            schemeColor73.Append(luminanceOffset30);

            solidFill35.Append(schemeColor73);
            A.Round round18 = new A.Round();

            outline29.Append(solidFill35);
            outline29.Append(round18);

            shapeProperties18.Append(outline29);

            seriesLine1.Append(lineReference24);
            seriesLine1.Append(fillReference24);
            seriesLine1.Append(effectReference24);
            seriesLine1.Append(fontReference24);
            seriesLine1.Append(shapeProperties18);

            Cs.TitleStyle titleStyle1 = new Cs.TitleStyle();
            Cs.LineReference lineReference25 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference25 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference25 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference25 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor74 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation43 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset31 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor74.Append(luminanceModulation43);
            schemeColor74.Append(luminanceOffset31);

            fontReference25.Append(schemeColor74);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType9 = new Cs.TextCharacterPropertiesType() { FontSize = 1400, Bold = false, Kerning = 1200, Spacing = 0, Baseline = 0 };

            titleStyle1.Append(lineReference25);
            titleStyle1.Append(fillReference25);
            titleStyle1.Append(effectReference25);
            titleStyle1.Append(fontReference25);
            titleStyle1.Append(textCharacterPropertiesType9);

            Cs.TrendlineStyle trendlineStyle1 = new Cs.TrendlineStyle();

            Cs.LineReference lineReference26 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor7 = new Cs.StyleColor() { Val = "auto" };

            lineReference26.Append(styleColor7);
            Cs.FillReference fillReference26 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference26 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference26 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor75 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference26.Append(schemeColor75);

            Cs.ShapeProperties shapeProperties19 = new Cs.ShapeProperties();

            A.Outline outline30 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill36 = new A.SolidFill();
            A.SchemeColor schemeColor76 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill36.Append(schemeColor76);
            A.PresetDash presetDash4 = new A.PresetDash() { Val = A.PresetLineDashValues.SystemDot };

            outline30.Append(solidFill36);
            outline30.Append(presetDash4);

            shapeProperties19.Append(outline30);

            trendlineStyle1.Append(lineReference26);
            trendlineStyle1.Append(fillReference26);
            trendlineStyle1.Append(effectReference26);
            trendlineStyle1.Append(fontReference26);
            trendlineStyle1.Append(shapeProperties19);

            Cs.TrendlineLabel trendlineLabel1 = new Cs.TrendlineLabel();
            Cs.LineReference lineReference27 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference27 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference27 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference27 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor77 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation44 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset32 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor77.Append(luminanceModulation44);
            schemeColor77.Append(luminanceOffset32);

            fontReference27.Append(schemeColor77);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType10 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            trendlineLabel1.Append(lineReference27);
            trendlineLabel1.Append(fillReference27);
            trendlineLabel1.Append(effectReference27);
            trendlineLabel1.Append(fontReference27);
            trendlineLabel1.Append(textCharacterPropertiesType10);

            Cs.UpBar upBar1 = new Cs.UpBar();
            Cs.LineReference lineReference28 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference28 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference28 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference28 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor78 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference28.Append(schemeColor78);

            Cs.ShapeProperties shapeProperties20 = new Cs.ShapeProperties();

            A.SolidFill solidFill37 = new A.SolidFill();
            A.SchemeColor schemeColor79 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            solidFill37.Append(schemeColor79);

            A.Outline outline31 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill38 = new A.SolidFill();

            A.SchemeColor schemeColor80 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation45 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset33 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor80.Append(luminanceModulation45);
            schemeColor80.Append(luminanceOffset33);

            solidFill38.Append(schemeColor80);
            A.Round round19 = new A.Round();

            outline31.Append(solidFill38);
            outline31.Append(round19);

            shapeProperties20.Append(solidFill37);
            shapeProperties20.Append(outline31);

            upBar1.Append(lineReference28);
            upBar1.Append(fillReference28);
            upBar1.Append(effectReference28);
            upBar1.Append(fontReference28);
            upBar1.Append(shapeProperties20);

            Cs.ValueAxis valueAxis3 = new Cs.ValueAxis();
            Cs.LineReference lineReference29 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference29 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference29 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference29 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor81 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation46 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset34 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor81.Append(luminanceModulation46);
            schemeColor81.Append(luminanceOffset34);

            fontReference29.Append(schemeColor81);

            Cs.ShapeProperties shapeProperties21 = new Cs.ShapeProperties();

            A.Outline outline32 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill39 = new A.SolidFill();

            A.SchemeColor schemeColor82 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation47 = new A.LuminanceModulation() { Val = 25000 };
            A.LuminanceOffset luminanceOffset35 = new A.LuminanceOffset() { Val = 75000 };

            schemeColor82.Append(luminanceModulation47);
            schemeColor82.Append(luminanceOffset35);

            solidFill39.Append(schemeColor82);
            A.Round round20 = new A.Round();

            outline32.Append(solidFill39);
            outline32.Append(round20);

            shapeProperties21.Append(outline32);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType11 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            valueAxis3.Append(lineReference29);
            valueAxis3.Append(fillReference29);
            valueAxis3.Append(effectReference29);
            valueAxis3.Append(fontReference29);
            valueAxis3.Append(shapeProperties21);
            valueAxis3.Append(textCharacterPropertiesType11);

            Cs.Wall wall1 = new Cs.Wall();
            Cs.LineReference lineReference30 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference30 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference30 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference30 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor83 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference30.Append(schemeColor83);

            Cs.ShapeProperties shapeProperties22 = new Cs.ShapeProperties();
            A.NoFill noFill13 = new A.NoFill();

            A.Outline outline33 = new A.Outline();
            A.NoFill noFill14 = new A.NoFill();

            outline33.Append(noFill14);

            shapeProperties22.Append(noFill13);
            shapeProperties22.Append(outline33);

            wall1.Append(lineReference30);
            wall1.Append(fillReference30);
            wall1.Append(effectReference30);
            wall1.Append(fontReference30);
            wall1.Append(shapeProperties22);

            chartStyle1.Append(axisTitle1);
            chartStyle1.Append(categoryAxis1);
            chartStyle1.Append(chartArea1);
            chartStyle1.Append(dataLabel1);
            chartStyle1.Append(dataLabelCallout1);
            chartStyle1.Append(dataPoint1);
            chartStyle1.Append(dataPoint3D1);
            chartStyle1.Append(dataPointLine1);
            chartStyle1.Append(dataPointMarker1);
            chartStyle1.Append(markerLayoutProperties1);
            chartStyle1.Append(dataPointWireframe1);
            chartStyle1.Append(dataTableStyle1);
            chartStyle1.Append(downBar1);
            chartStyle1.Append(dropLine1);
            chartStyle1.Append(errorBar1);
            chartStyle1.Append(floor1);
            chartStyle1.Append(gridlineMajor1);
            chartStyle1.Append(gridlineMinor1);
            chartStyle1.Append(hiLoLine1);
            chartStyle1.Append(leaderLine1);
            chartStyle1.Append(legendStyle1);
            chartStyle1.Append(plotArea2);
            chartStyle1.Append(plotArea3D1);
            chartStyle1.Append(seriesAxis1);
            chartStyle1.Append(seriesLine1);
            chartStyle1.Append(titleStyle1);
            chartStyle1.Append(trendlineStyle1);
            chartStyle1.Append(trendlineLabel1);
            chartStyle1.Append(upBar1);
            chartStyle1.Append(valueAxis3);
            chartStyle1.Append(wall1);

            chartStylePart1.ChartStyle = chartStyle1;
        }

        // Generates content of spreadsheetPrinterSettingsPart1.
        private void GenerateSpreadsheetPrinterSettingsPart1Content(SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(spreadsheetPrinterSettingsPart1Data);
            spreadsheetPrinterSettingsPart1.FeedData(data);
            data.Close();
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

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Josue Garcia";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2017-03-21T16:52:52Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2017-04-04T00:11:15Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Josue Garcia";
        }

        #region Binary Data
        private string spreadsheetPrinterSettingsPart1Data = "TQBpAGMAcgBvAHMAbwBmAHQAIABQAHIAaQBuAHQAIAB0AG8AIABQAEQARgAAAAAAAAAAAAAAAAAAAAAAAAAAAAEEAwbcAFAUAy8BAAEAAQDqCm8IZAABAA8AWAICAAEAWAIDAAEATABlAHQAdABlAHIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAgAAAAEAAAD/////R0lTNAAAAAAAAAAAAAAAAERJTlUiAMgAJAMsET9de34AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABQAAAAAACQABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADIAAAAU01USgAAAAAQALgAewAwADgANABGADAAMQBGAEEALQBFADYAMwA0AC0ANABEADcANwAtADgAMwBFAEUALQAwADcANAA4ADEANwBDADAAMwA1ADgAMQB9AAAAUkVTRExMAFVuaXJlc0RMTABQYXBlclNpemUATEVUVEVSAE9yaWVudGF0aW9uAFBPUlRSQUlUAFJlc29sdXRpb24AUmVzT3B0aW9uMQBDb2xvck1vZGUAQ29sb3IAAAAAAAAAAAAAAAAAACwRAABWNERNAQAAAAAAAACcCnAiHAAAAOwAAAADAAAA+gFPCDTmd02D7gdIF8A1gdAAAABMAAAAAwAAAAAIAAAAAAAAAAAAAAMAAAAACAAAKgAAAAAIAAADAAAAQAAAAFYAAAAAEAAARABvAGMAdQBtAGUAbgB0AFUAcwBlAHIAUABhAHMAcwB3AG8AcgBkAAAARABvAGMAdQBtAGUAbgB0AE8AdwBuAGUAcgBQAGEAcwBzAHcAbwByAGQAAABEAG8AYwB1AG0AZQBuAHQAQwByAHkAcAB0AFMAZQBjAHUAcgBpAHQAeQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

    }
}
