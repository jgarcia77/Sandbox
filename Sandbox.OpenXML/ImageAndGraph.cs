﻿using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;

namespace GeneratedCode
{
    public class GeneratedClass
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

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPart1Content(workbookStylesPart1);

            ThemePart themePart1 = workbookPart1.AddNewPart<ThemePart>("rId2");
            GenerateThemePart1Content(themePart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPart1Content(worksheetPart1);

            DrawingsPart drawingsPart1 = worksheetPart1.AddNewPart<DrawingsPart>("rId2");
            GenerateDrawingsPart1Content(drawingsPart1);

            ChartPart chartPart1 = drawingsPart1.AddNewPart<ChartPart>("rId2");
            GenerateChartPart1Content(chartPart1);

            ImagePart imagePart1 = drawingsPart1.AddNewPart<ImagePart>("image/tiff", "rId1");
            GenerateImagePart1Content(imagePart1);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1 = worksheetPart1.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart1Content(spreadsheetPrinterSettingsPart1);

            SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId4");
            GenerateSharedStringTablePart1Content(sharedStringTablePart1);

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

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Worksheets";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Results Report";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
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
            WorkbookProperties workbookProperties1 = new WorkbookProperties();

            AlternateContent alternateContent1 = new AlternateContent();
            alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "x15" };

            X15ac.AbsolutePath absolutePath1 = new X15ac.AbsolutePath() { Url = "C:\\Users\\josueg\\Documents\\Projects\\Deloitte\\STAR Modernization - Deloitte Documents\\" };
            absolutePath1.AddNamespaceDeclaration("x15ac", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac");

            alternateContentChoice1.Append(absolutePath1);

            alternateContent1.Append(alternateContentChoice1);

            BookViews bookViews1 = new BookViews();
            WorkbookView workbookView1 = new WorkbookView() { XWindow = 0, YWindow = 460, WindowWidth = (UInt32Value)28160U, WindowHeight = (UInt32Value)16880U, TabRatio = (UInt32Value)500U };

            bookViews1.Append(workbookView1);

            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Results Report", SheetId = (UInt32Value)1U, Id = "rId1" };

            sheets1.Append(sheet1);
            CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)150000U, CalculationOnSave = false };

            WorkbookExtensionList workbookExtensionList1 = new WorkbookExtensionList();

            WorkbookExtension workbookExtension1 = new WorkbookExtension() { Uri = "{7523E5D3-25F3-A5E0-1632-64F254C22452}" };
            workbookExtension1.AddNamespaceDeclaration("mx", "http://schemas.microsoft.com/office/mac/excel/2008/main");

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<mx:ArchID Flags=\"2\" xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\" />");

            workbookExtension1.Append(openXmlUnknownElement1);

            workbookExtensionList1.Append(workbookExtension1);

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(alternateContent1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets1);
            workbook1.Append(calculationProperties1);
            workbook1.Append(workbookExtensionList1);

            workbookPart1.Workbook = workbook1;
        }

        // Generates content of workbookStylesPart1.
        private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)4U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 12D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            FontSize fontSize2 = new FontSize() { Val = 8D };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(fontSize2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            Font font3 = new Font();
            Underline underline1 = new Underline();
            FontSize fontSize3 = new FontSize() { Val = 12D };
            Color color2 = new Color() { Theme = (UInt32Value)10U };
            FontName fontName3 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

            font3.Append(underline1);
            font3.Append(fontSize3);
            font3.Append(color2);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);
            font3.Append(fontScheme3);

            Font font4 = new Font();
            Underline underline2 = new Underline();
            FontSize fontSize4 = new FontSize() { Val = 12D };
            Color color3 = new Color() { Theme = (UInt32Value)11U };
            FontName fontName4 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme4 = new FontScheme() { Val = FontSchemeValues.Minor };

            font4.Append(underline2);
            font4.Append(fontSize4);
            font4.Append(color3);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering4);
            font4.Append(fontScheme4);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

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

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)3U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };

            cellStyleFormats1.Append(cellFormat1);
            cellStyleFormats1.Append(cellFormat2);
            cellStyleFormats1.Append(cellFormat3);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)2U };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };

            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)3U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)2U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle2 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)1U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle3 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            cellStyles1.Append(cellStyle2);
            cellStyles1.Append(cellStyle3);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium9", DefaultPivotStyle = "PivotStyleMedium7" };

            Colors colors1 = new Colors();

            MruColors mruColors1 = new MruColors();
            Color color4 = new Color() { Rgb = "FF0076A8" };
            Color color5 = new Color() { Rgb = "FFD0D2D3" };
            Color color6 = new Color() { Rgb = "FFF5F7F8" };
            Color color7 = new Color() { Rgb = "FFF4F7F8" };
            Color color8 = new Color() { Rgb = "FFFFCCCC" };

            mruColors1.Append(color4);
            mruColors1.Append(color5);
            mruColors1.Append(color6);
            mruColors1.Append(color7);
            mruColors1.Append(color8);

            colors1.Append(mruColors1);

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

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(colors1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
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
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4472C4" };

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
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "5B9BD5" };

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

            A.FontScheme fontScheme5 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "Yu Gothic Light" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "DengXian Light" };
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
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "Yu Gothic" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "DengXian" };
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

            fontScheme5.Append(majorFont1);
            fontScheme5.Append(minorFont1);

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
            themeElements1.Append(fontScheme5);
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
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "B1" };

            SheetViews sheetViews1 = new SheetViews();
            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultColumnWidth = 10.6640625D, DefaultRowHeight = 15.5D, DyDescent = 0.35D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 4D, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 9.1640625D, Style = (UInt32Value)1U, BestFit = true, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)7U, Width = 13.33203125D, CustomWidth = true };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            SheetData sheetData1 = new SheetData();
            PhoneticProperties phoneticProperties1 = new PhoneticProperties() { FontId = (UInt32Value)1U, Type = PhoneticValues.NoConversion };
            PrintOptions printOptions1 = new PrintOptions() { HorizontalCentered = true };
            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup() { Scale = (UInt32Value)76U, Orientation = OrientationValues.Portrait, HorizontalDpi = (UInt32Value)200U, VerticalDpi = (UInt32Value)200U, Id = "rId1" };
            Drawing drawing1 = new Drawing() { Id = "rId2" };

            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(phoneticProperties1);
            worksheet1.Append(printOptions1);
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

            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };

            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            columnId1.Text = "0";
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            columnOffset1.Text = "0";
            Xdr.RowId rowId1 = new Xdr.RowId();
            rowId1.Text = "0";
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            rowOffset1.Text = "18904";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            columnId2.Text = "2";
            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            columnOffset2.Text = "824661";
            Xdr.RowId rowId2 = new Xdr.RowId();
            rowId2.Text = "2";
            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            rowOffset2.Text = "62530";

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            Xdr.Picture picture1 = new Xdr.Picture();

            Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new Xdr.NonVisualPictureProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture 3" };

            Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Xdr.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Xdr.BlipFill blipFill1 = new Xdr.BlipFill() { RotateWithShape = true };

            A.Blip blip1 = new A.Blip() { Embed = "rId1" };
            blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            A.SourceRectangle sourceRectangle1 = new A.SourceRectangle() { Top = 32250, Bottom = 34913 };
            A.Stretch stretch1 = new A.Stretch();

            blipFill1.Append(blip1);
            blipFill1.Append(sourceRectangle1);
            blipFill1.Append(stretch1);

            Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 18904L };
            A.Extents extents1 = new A.Extents() { Cx = 1827961L, Cy = 437326L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(picture1);
            twoCellAnchor1.Append(clientData1);

            Xdr.TwoCellAnchor twoCellAnchor2 = new Xdr.TwoCellAnchor();

            Xdr.FromMarker fromMarker2 = new Xdr.FromMarker();
            Xdr.ColumnId columnId3 = new Xdr.ColumnId();
            columnId3.Text = "0";
            Xdr.ColumnOffset columnOffset3 = new Xdr.ColumnOffset();
            columnOffset3.Text = "0";
            Xdr.RowId rowId3 = new Xdr.RowId();
            rowId3.Text = "3";
            Xdr.RowOffset rowOffset3 = new Xdr.RowOffset();
            rowOffset3.Text = "0";

            fromMarker2.Append(columnId3);
            fromMarker2.Append(columnOffset3);
            fromMarker2.Append(rowId3);
            fromMarker2.Append(rowOffset3);

            Xdr.ToMarker toMarker2 = new Xdr.ToMarker();
            Xdr.ColumnId columnId4 = new Xdr.ColumnId();
            columnId4.Text = "4";
            Xdr.ColumnOffset columnOffset4 = new Xdr.ColumnOffset();
            columnOffset4.Text = "799194";
            Xdr.RowId rowId4 = new Xdr.RowId();
            rowId4.Text = "36";
            Xdr.RowOffset rowOffset4 = new Xdr.RowOffset();
            rowOffset4.Text = "48192";

            toMarker2.Append(columnId4);
            toMarker2.Append(columnOffset4);
            toMarker2.Append(rowId4);
            toMarker2.Append(rowOffset4);

            Xdr.GraphicFrame graphicFrame1 = new Xdr.GraphicFrame() { Macro = "" };

            Xdr.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = new Xdr.NonVisualGraphicFrameProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Residuals" };

            Xdr.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Xdr.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            nonVisualGraphicFrameProperties1.Append(nonVisualDrawingProperties2);
            nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameDrawingProperties1);

            Xdr.Transform transform1 = new Xdr.Transform();
            A.Offset offset2 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents2 = new A.Extents() { Cx = 0L, Cy = 0L };

            transform1.Append(offset2);
            transform1.Append(extents2);

            A.Graphic graphic1 = new A.Graphic();

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            C.ChartReference chartReference1 = new C.ChartReference() { Id = "rId2" };
            chartReference1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartReference1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            graphicData1.Append(chartReference1);

            graphic1.Append(graphicData1);

            graphicFrame1.Append(nonVisualGraphicFrameProperties1);
            graphicFrame1.Append(transform1);
            graphicFrame1.Append(graphic1);
            Xdr.ClientData clientData2 = new Xdr.ClientData();

            twoCellAnchor2.Append(fromMarker2);
            twoCellAnchor2.Append(toMarker2);
            twoCellAnchor2.Append(graphicFrame1);
            twoCellAnchor2.Append(clientData2);

            worksheetDrawing1.Append(twoCellAnchor1);
            worksheetDrawing1.Append(twoCellAnchor2);

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
            C.RoundedCorners roundedCorners1 = new C.RoundedCorners() { Val = true };

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

            C.ScatterChart scatterChart1 = new C.ScatterChart();
            C.ScatterStyle scatterStyle1 = new C.ScatterStyle() { Val = C.ScatterStyleValues.SmoothMarker };
            C.VaryColors varyColors1 = new C.VaryColors() { Val = false };

            C.ScatterChartSeries scatterChartSeries1 = new C.ScatterChartSeries();
            C.Index index1 = new C.Index() { Val = (UInt32Value)0U };
            C.Order order1 = new C.Order() { Val = (UInt32Value)0U };

            C.Marker marker1 = new C.Marker();
            C.Symbol symbol1 = new C.Symbol() { Val = C.MarkerStyleValues.Circle };
            C.Size size1 = new C.Size() { Val = 5 };

            marker1.Append(symbol1);
            marker1.Append(size1);

            C.XValues xValues1 = new C.XValues();

            C.NumberReference numberReference1 = new C.NumberReference();
            C.Formula formula1 = new C.Formula();
            formula1.Text = "E$11:E46";

            C.NumberingCache numberingCache1 = new C.NumberingCache();
            C.FormatCode formatCode1 = new C.FormatCode();
            formatCode1.Text = "General";
            C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)36U };

            C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue1 = new C.NumericValue();
            numericValue1.Text = "-1.3919999999999999";

            numericPoint1.Append(numericValue1);

            C.NumericPoint numericPoint2 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue2 = new C.NumericValue();
            numericValue2.Text = "-0.483333333333333";

            numericPoint2.Append(numericValue2);

            C.NumericPoint numericPoint3 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue3 = new C.NumericValue();
            numericValue3.Text = "1.871";

            numericPoint3.Append(numericValue3);

            C.NumericPoint numericPoint4 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue4 = new C.NumericValue();
            numericValue4.Text = "0.115333333333333";

            numericPoint4.Append(numericValue4);

            C.NumericPoint numericPoint5 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue5 = new C.NumericValue();
            numericValue5.Text = "-0.55033333333333301";

            numericPoint5.Append(numericValue5);

            C.NumericPoint numericPoint6 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue6 = new C.NumericValue();
            numericValue6.Text = "-2.19966666666667";

            numericPoint6.Append(numericValue6);

            C.NumericPoint numericPoint7 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue7 = new C.NumericValue();
            numericValue7.Text = "-0.13200000000000001";

            numericPoint7.Append(numericValue7);

            C.NumericPoint numericPoint8 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue8 = new C.NumericValue();
            numericValue8.Text = "-1.17766666666667";

            numericPoint8.Append(numericValue8);

            C.NumericPoint numericPoint9 = new C.NumericPoint() { Index = (UInt32Value)8U };
            C.NumericValue numericValue9 = new C.NumericValue();
            numericValue9.Text = "1.3480000000000001";

            numericPoint9.Append(numericValue9);

            C.NumericPoint numericPoint10 = new C.NumericPoint() { Index = (UInt32Value)9U };
            C.NumericValue numericValue10 = new C.NumericValue();
            numericValue10.Text = "-0.13100000000000001";

            numericPoint10.Append(numericValue10);

            C.NumericPoint numericPoint11 = new C.NumericPoint() { Index = (UInt32Value)10U };
            C.NumericValue numericValue11 = new C.NumericValue();
            numericValue11.Text = "0.62";

            numericPoint11.Append(numericValue11);

            C.NumericPoint numericPoint12 = new C.NumericPoint() { Index = (UInt32Value)11U };
            C.NumericValue numericValue12 = new C.NumericValue();
            numericValue12.Text = "0.628";

            numericPoint12.Append(numericValue12);

            C.NumericPoint numericPoint13 = new C.NumericPoint() { Index = (UInt32Value)12U };
            C.NumericValue numericValue13 = new C.NumericValue();
            numericValue13.Text = "-1.411999999999999";

            numericPoint13.Append(numericValue13);

            C.NumericPoint numericPoint14 = new C.NumericPoint() { Index = (UInt32Value)13U };
            C.NumericValue numericValue14 = new C.NumericValue();
            numericValue14.Text = "-2.3333333333333301E-3";

            numericPoint14.Append(numericValue14);

            C.NumericPoint numericPoint15 = new C.NumericPoint() { Index = (UInt32Value)14U };
            C.NumericValue numericValue15 = new C.NumericValue();
            numericValue15.Text = "-0.44600000000000001";

            numericPoint15.Append(numericValue15);

            C.NumericPoint numericPoint16 = new C.NumericPoint() { Index = (UInt32Value)15U };
            C.NumericValue numericValue16 = new C.NumericValue();
            numericValue16.Text = "1.0353333333333301";

            numericPoint16.Append(numericValue16);

            C.NumericPoint numericPoint17 = new C.NumericPoint() { Index = (UInt32Value)16U };
            C.NumericValue numericValue17 = new C.NumericValue();
            numericValue17.Text = "1.33033333333333";

            numericPoint17.Append(numericValue17);

            C.NumericPoint numericPoint18 = new C.NumericPoint() { Index = (UInt32Value)17U };
            C.NumericValue numericValue18 = new C.NumericValue();
            numericValue18.Text = "0.14499999999999999";

            numericPoint18.Append(numericValue18);

            C.NumericPoint numericPoint19 = new C.NumericPoint() { Index = (UInt32Value)18U };
            C.NumericValue numericValue19 = new C.NumericValue();
            numericValue19.Text = "2.5666666666666699E-2";

            numericPoint19.Append(numericValue19);

            C.NumericPoint numericPoint20 = new C.NumericPoint() { Index = (UInt32Value)19U };
            C.NumericValue numericValue20 = new C.NumericValue();
            numericValue20.Text = "-0.43333333333333302";

            numericPoint20.Append(numericValue20);

            C.NumericPoint numericPoint21 = new C.NumericPoint() { Index = (UInt32Value)20U };
            C.NumericValue numericValue21 = new C.NumericValue();
            numericValue21.Text = "0.664333333333333";

            numericPoint21.Append(numericValue21);

            C.NumericPoint numericPoint22 = new C.NumericPoint() { Index = (UInt32Value)21U };
            C.NumericValue numericValue22 = new C.NumericValue();
            numericValue22.Text = "-0.31766666666666699";

            numericPoint22.Append(numericValue22);

            C.NumericPoint numericPoint23 = new C.NumericPoint() { Index = (UInt32Value)22U };
            C.NumericValue numericValue23 = new C.NumericValue();
            numericValue23.Text = "1.1439999999999999";

            numericPoint23.Append(numericValue23);

            C.NumericPoint numericPoint24 = new C.NumericPoint() { Index = (UInt32Value)23U };
            C.NumericValue numericValue24 = new C.NumericValue();
            numericValue24.Text = "-0.24966666666666701";

            numericPoint24.Append(numericValue24);

            C.NumericPoint numericPoint25 = new C.NumericPoint() { Index = (UInt32Value)24U };
            C.NumericValue numericValue25 = new C.NumericValue();
            numericValue25.Text = "-0.174666666666667";

            numericPoint25.Append(numericValue25);

            C.NumericPoint numericPoint26 = new C.NumericPoint() { Index = (UInt32Value)25U };
            C.NumericValue numericValue26 = new C.NumericValue();
            numericValue26.Text = "0.73799999999999999";

            numericPoint26.Append(numericValue26);

            C.NumericPoint numericPoint27 = new C.NumericPoint() { Index = (UInt32Value)26U };
            C.NumericValue numericValue27 = new C.NumericValue();
            numericValue27.Text = "0.98633333333333295";

            numericPoint27.Append(numericValue27);

            C.NumericPoint numericPoint28 = new C.NumericPoint() { Index = (UInt32Value)27U };
            C.NumericValue numericValue28 = new C.NumericValue();
            numericValue28.Text = "1.2896666666666701";

            numericPoint28.Append(numericValue28);

            C.NumericPoint numericPoint29 = new C.NumericPoint() { Index = (UInt32Value)28U };
            C.NumericValue numericValue29 = new C.NumericValue();
            numericValue29.Text = "0.75900000000000001";

            numericPoint29.Append(numericValue29);

            C.NumericPoint numericPoint30 = new C.NumericPoint() { Index = (UInt32Value)29U };
            C.NumericValue numericValue30 = new C.NumericValue();
            numericValue30.Text = "-0.65200000000000002";

            numericPoint30.Append(numericValue30);

            C.NumericPoint numericPoint31 = new C.NumericPoint() { Index = (UInt32Value)30U };
            C.NumericValue numericValue31 = new C.NumericValue();
            numericValue31.Text = "0.26233333333333297";

            numericPoint31.Append(numericValue31);

            C.NumericPoint numericPoint32 = new C.NumericPoint() { Index = (UInt32Value)31U };
            C.NumericValue numericValue32 = new C.NumericValue();
            numericValue32.Text = "-0.64733333333333298";

            numericPoint32.Append(numericValue32);

            C.NumericPoint numericPoint33 = new C.NumericPoint() { Index = (UInt32Value)32U };
            C.NumericValue numericValue33 = new C.NumericValue();
            numericValue33.Text = "1.98";

            numericPoint33.Append(numericValue33);

            C.NumericPoint numericPoint34 = new C.NumericPoint() { Index = (UInt32Value)33U };
            C.NumericValue numericValue34 = new C.NumericValue();
            numericValue34.Text = "1.798";

            numericPoint34.Append(numericValue34);

            C.NumericPoint numericPoint35 = new C.NumericPoint() { Index = (UInt32Value)34U };
            C.NumericValue numericValue35 = new C.NumericValue();
            numericValue35.Text = "0.77233333333333298";

            numericPoint35.Append(numericValue35);

            C.NumericPoint numericPoint36 = new C.NumericPoint() { Index = (UInt32Value)35U };
            C.NumericValue numericValue36 = new C.NumericValue();
            numericValue36.Text = "-0.54500000000000004";

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
            formula2.Text = "A$11:A46";

            C.NumberingCache numberingCache2 = new C.NumberingCache();
            C.FormatCode formatCode2 = new C.FormatCode();
            formatCode2.Text = "General";
            C.PointCount pointCount2 = new C.PointCount() { Val = (UInt32Value)36U };

            C.NumericPoint numericPoint37 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue37 = new C.NumericValue();
            numericValue37.Text = "1";

            numericPoint37.Append(numericValue37);

            C.NumericPoint numericPoint38 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue38 = new C.NumericValue();
            numericValue38.Text = "2";

            numericPoint38.Append(numericValue38);

            C.NumericPoint numericPoint39 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue39 = new C.NumericValue();
            numericValue39.Text = "3";

            numericPoint39.Append(numericValue39);

            C.NumericPoint numericPoint40 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue40 = new C.NumericValue();
            numericValue40.Text = "4";

            numericPoint40.Append(numericValue40);

            C.NumericPoint numericPoint41 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue41 = new C.NumericValue();
            numericValue41.Text = "5";

            numericPoint41.Append(numericValue41);

            C.NumericPoint numericPoint42 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue42 = new C.NumericValue();
            numericValue42.Text = "6";

            numericPoint42.Append(numericValue42);

            C.NumericPoint numericPoint43 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue43 = new C.NumericValue();
            numericValue43.Text = "7";

            numericPoint43.Append(numericValue43);

            C.NumericPoint numericPoint44 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue44 = new C.NumericValue();
            numericValue44.Text = "8";

            numericPoint44.Append(numericValue44);

            C.NumericPoint numericPoint45 = new C.NumericPoint() { Index = (UInt32Value)8U };
            C.NumericValue numericValue45 = new C.NumericValue();
            numericValue45.Text = "9";

            numericPoint45.Append(numericValue45);

            C.NumericPoint numericPoint46 = new C.NumericPoint() { Index = (UInt32Value)9U };
            C.NumericValue numericValue46 = new C.NumericValue();
            numericValue46.Text = "10";

            numericPoint46.Append(numericValue46);

            C.NumericPoint numericPoint47 = new C.NumericPoint() { Index = (UInt32Value)10U };
            C.NumericValue numericValue47 = new C.NumericValue();
            numericValue47.Text = "11";

            numericPoint47.Append(numericValue47);

            C.NumericPoint numericPoint48 = new C.NumericPoint() { Index = (UInt32Value)11U };
            C.NumericValue numericValue48 = new C.NumericValue();
            numericValue48.Text = "12";

            numericPoint48.Append(numericValue48);

            C.NumericPoint numericPoint49 = new C.NumericPoint() { Index = (UInt32Value)12U };
            C.NumericValue numericValue49 = new C.NumericValue();
            numericValue49.Text = "13";

            numericPoint49.Append(numericValue49);

            C.NumericPoint numericPoint50 = new C.NumericPoint() { Index = (UInt32Value)13U };
            C.NumericValue numericValue50 = new C.NumericValue();
            numericValue50.Text = "14";

            numericPoint50.Append(numericValue50);

            C.NumericPoint numericPoint51 = new C.NumericPoint() { Index = (UInt32Value)14U };
            C.NumericValue numericValue51 = new C.NumericValue();
            numericValue51.Text = "15";

            numericPoint51.Append(numericValue51);

            C.NumericPoint numericPoint52 = new C.NumericPoint() { Index = (UInt32Value)15U };
            C.NumericValue numericValue52 = new C.NumericValue();
            numericValue52.Text = "16";

            numericPoint52.Append(numericValue52);

            C.NumericPoint numericPoint53 = new C.NumericPoint() { Index = (UInt32Value)16U };
            C.NumericValue numericValue53 = new C.NumericValue();
            numericValue53.Text = "17";

            numericPoint53.Append(numericValue53);

            C.NumericPoint numericPoint54 = new C.NumericPoint() { Index = (UInt32Value)17U };
            C.NumericValue numericValue54 = new C.NumericValue();
            numericValue54.Text = "18";

            numericPoint54.Append(numericValue54);

            C.NumericPoint numericPoint55 = new C.NumericPoint() { Index = (UInt32Value)18U };
            C.NumericValue numericValue55 = new C.NumericValue();
            numericValue55.Text = "19";

            numericPoint55.Append(numericValue55);

            C.NumericPoint numericPoint56 = new C.NumericPoint() { Index = (UInt32Value)19U };
            C.NumericValue numericValue56 = new C.NumericValue();
            numericValue56.Text = "20";

            numericPoint56.Append(numericValue56);

            C.NumericPoint numericPoint57 = new C.NumericPoint() { Index = (UInt32Value)20U };
            C.NumericValue numericValue57 = new C.NumericValue();
            numericValue57.Text = "21";

            numericPoint57.Append(numericValue57);

            C.NumericPoint numericPoint58 = new C.NumericPoint() { Index = (UInt32Value)21U };
            C.NumericValue numericValue58 = new C.NumericValue();
            numericValue58.Text = "22";

            numericPoint58.Append(numericValue58);

            C.NumericPoint numericPoint59 = new C.NumericPoint() { Index = (UInt32Value)22U };
            C.NumericValue numericValue59 = new C.NumericValue();
            numericValue59.Text = "23";

            numericPoint59.Append(numericValue59);

            C.NumericPoint numericPoint60 = new C.NumericPoint() { Index = (UInt32Value)23U };
            C.NumericValue numericValue60 = new C.NumericValue();
            numericValue60.Text = "24";

            numericPoint60.Append(numericValue60);

            C.NumericPoint numericPoint61 = new C.NumericPoint() { Index = (UInt32Value)24U };
            C.NumericValue numericValue61 = new C.NumericValue();
            numericValue61.Text = "25";

            numericPoint61.Append(numericValue61);

            C.NumericPoint numericPoint62 = new C.NumericPoint() { Index = (UInt32Value)25U };
            C.NumericValue numericValue62 = new C.NumericValue();
            numericValue62.Text = "26";

            numericPoint62.Append(numericValue62);

            C.NumericPoint numericPoint63 = new C.NumericPoint() { Index = (UInt32Value)26U };
            C.NumericValue numericValue63 = new C.NumericValue();
            numericValue63.Text = "27";

            numericPoint63.Append(numericValue63);

            C.NumericPoint numericPoint64 = new C.NumericPoint() { Index = (UInt32Value)27U };
            C.NumericValue numericValue64 = new C.NumericValue();
            numericValue64.Text = "28";

            numericPoint64.Append(numericValue64);

            C.NumericPoint numericPoint65 = new C.NumericPoint() { Index = (UInt32Value)28U };
            C.NumericValue numericValue65 = new C.NumericValue();
            numericValue65.Text = "29";

            numericPoint65.Append(numericValue65);

            C.NumericPoint numericPoint66 = new C.NumericPoint() { Index = (UInt32Value)29U };
            C.NumericValue numericValue66 = new C.NumericValue();
            numericValue66.Text = "30";

            numericPoint66.Append(numericValue66);

            C.NumericPoint numericPoint67 = new C.NumericPoint() { Index = (UInt32Value)30U };
            C.NumericValue numericValue67 = new C.NumericValue();
            numericValue67.Text = "31";

            numericPoint67.Append(numericValue67);

            C.NumericPoint numericPoint68 = new C.NumericPoint() { Index = (UInt32Value)31U };
            C.NumericValue numericValue68 = new C.NumericValue();
            numericValue68.Text = "32";

            numericPoint68.Append(numericValue68);

            C.NumericPoint numericPoint69 = new C.NumericPoint() { Index = (UInt32Value)32U };
            C.NumericValue numericValue69 = new C.NumericValue();
            numericValue69.Text = "33";

            numericPoint69.Append(numericValue69);

            C.NumericPoint numericPoint70 = new C.NumericPoint() { Index = (UInt32Value)33U };
            C.NumericValue numericValue70 = new C.NumericValue();
            numericValue70.Text = "34";

            numericPoint70.Append(numericValue70);

            C.NumericPoint numericPoint71 = new C.NumericPoint() { Index = (UInt32Value)34U };
            C.NumericValue numericValue71 = new C.NumericValue();
            numericValue71.Text = "35";

            numericPoint71.Append(numericValue71);

            C.NumericPoint numericPoint72 = new C.NumericPoint() { Index = (UInt32Value)35U };
            C.NumericValue numericValue72 = new C.NumericValue();
            numericValue72.Text = "36";

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
            C.Smooth smooth1 = new C.Smooth() { Val = true };

            C.ScatterSerExtensionList scatterSerExtensionList1 = new C.ScatterSerExtensionList();
            scatterSerExtensionList1.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");

            C.ScatterSerExtension scatterSerExtension1 = new C.ScatterSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
            scatterSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000000-0324-4D3A-B691-0B4389CE1DE7}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

            scatterSerExtension1.Append(openXmlUnknownElement2);

            scatterSerExtensionList1.Append(scatterSerExtension1);

            scatterChartSeries1.Append(index1);
            scatterChartSeries1.Append(order1);
            scatterChartSeries1.Append(marker1);
            scatterChartSeries1.Append(xValues1);
            scatterChartSeries1.Append(yValues1);
            scatterChartSeries1.Append(smooth1);
            scatterChartSeries1.Append(scatterSerExtensionList1);

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
            C.AxisId axisId1 = new C.AxisId() { Val = (UInt32Value)428045344U };
            C.AxisId axisId2 = new C.AxisId() { Val = (UInt32Value)428047304U };

            scatterChart1.Append(scatterStyle1);
            scatterChart1.Append(varyColors1);
            scatterChart1.Append(scatterChartSeries1);
            scatterChart1.Append(dataLabels1);
            scatterChart1.Append(axisId1);
            scatterChart1.Append(axisId2);

            C.ValueAxis valueAxis1 = new C.ValueAxis();
            C.AxisId axisId3 = new C.AxisId() { Val = (UInt32Value)428045344U };

            C.Scaling scaling1 = new C.Scaling();
            C.Orientation orientation1 = new C.Orientation() { Val = C.OrientationValues.MinMax };
            C.MaxAxisValue maxAxisValue1 = new C.MaxAxisValue() { Val = 4D };
            C.MinAxisValue minAxisValue1 = new C.MinAxisValue() { Val = -4D };

            scaling1.Append(orientation1);
            scaling1.Append(maxAxisValue1);
            scaling1.Append(minAxisValue1);
            C.Delete delete1 = new C.Delete() { Val = false };
            C.AxisPosition axisPosition1 = new C.AxisPosition() { Val = C.AxisPositionValues.Top };

            C.MajorGridlines majorGridlines1 = new C.MajorGridlines();

            C.ChartShapeProperties chartShapeProperties1 = new C.ChartShapeProperties();

            A.Outline outline4 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill7 = new A.SolidFill();

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation9 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor16.Append(luminanceModulation9);
            schemeColor16.Append(luminanceOffset1);

            solidFill7.Append(schemeColor16);
            A.Round round1 = new A.Round();

            outline4.Append(solidFill7);
            outline4.Append(round1);
            A.EffectList effectList4 = new A.EffectList();

            chartShapeProperties1.Append(outline4);
            chartShapeProperties1.Append(effectList4);

            majorGridlines1.Append(chartShapeProperties1);
            C.NumberingFormat numberingFormat1 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
            C.MajorTickMark majorTickMark1 = new C.MajorTickMark() { Val = C.TickMarkValues.Outside };
            C.MinorTickMark minorTickMark1 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
            C.TickLabelPosition tickLabelPosition1 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.Low };

            C.ChartShapeProperties chartShapeProperties2 = new C.ChartShapeProperties();
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline5 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill8 = new A.SolidFill();

            A.SchemeColor schemeColor17 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation10 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor17.Append(luminanceModulation10);
            schemeColor17.Append(luminanceOffset2);

            solidFill8.Append(schemeColor17);
            A.Round round2 = new A.Round();

            outline5.Append(solidFill8);
            outline5.Append(round2);
            A.EffectList effectList5 = new A.EffectList();

            chartShapeProperties2.Append(noFill1);
            chartShapeProperties2.Append(outline5);
            chartShapeProperties2.Append(effectList5);

            C.TextProperties textProperties1 = new C.TextProperties();
            A.BodyProperties bodyProperties1 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill9 = new A.SolidFill();

            A.SchemeColor schemeColor18 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation11 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor18.Append(luminanceModulation11);
            schemeColor18.Append(luminanceOffset3);

            solidFill9.Append(schemeColor18);
            A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont3 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties1.Append(solidFill9);
            defaultRunProperties1.Append(latinFont3);
            defaultRunProperties1.Append(eastAsianFont3);
            defaultRunProperties1.Append(complexScriptFont3);

            paragraphProperties1.Append(defaultRunProperties1);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(endParagraphRunProperties1);

            textProperties1.Append(bodyProperties1);
            textProperties1.Append(listStyle1);
            textProperties1.Append(paragraph1);
            C.CrossingAxis crossingAxis1 = new C.CrossingAxis() { Val = (UInt32Value)428047304U };
            C.Crosses crosses1 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
            C.CrossBetween crossBetween1 = new C.CrossBetween() { Val = C.CrossBetweenValues.MidpointCategory };
            C.MajorUnit majorUnit1 = new C.MajorUnit() { Val = 1D };
            C.MinorUnit minorUnit1 = new C.MinorUnit() { Val = 1D };

            valueAxis1.Append(axisId3);
            valueAxis1.Append(scaling1);
            valueAxis1.Append(delete1);
            valueAxis1.Append(axisPosition1);
            valueAxis1.Append(majorGridlines1);
            valueAxis1.Append(numberingFormat1);
            valueAxis1.Append(majorTickMark1);
            valueAxis1.Append(minorTickMark1);
            valueAxis1.Append(tickLabelPosition1);
            valueAxis1.Append(chartShapeProperties2);
            valueAxis1.Append(textProperties1);
            valueAxis1.Append(crossingAxis1);
            valueAxis1.Append(crosses1);
            valueAxis1.Append(crossBetween1);
            valueAxis1.Append(majorUnit1);
            valueAxis1.Append(minorUnit1);

            C.ValueAxis valueAxis2 = new C.ValueAxis();
            C.AxisId axisId4 = new C.AxisId() { Val = (UInt32Value)428047304U };

            C.Scaling scaling2 = new C.Scaling();
            C.Orientation orientation2 = new C.Orientation() { Val = C.OrientationValues.MaxMin };
            C.MaxAxisValue maxAxisValue2 = new C.MaxAxisValue() { Val = 36D };
            C.MinAxisValue minAxisValue2 = new C.MinAxisValue() { Val = 1D };

            scaling2.Append(orientation2);
            scaling2.Append(maxAxisValue2);
            scaling2.Append(minAxisValue2);
            C.Delete delete2 = new C.Delete() { Val = false };
            C.AxisPosition axisPosition2 = new C.AxisPosition() { Val = C.AxisPositionValues.Left };

            C.MajorGridlines majorGridlines2 = new C.MajorGridlines();

            C.ChartShapeProperties chartShapeProperties3 = new C.ChartShapeProperties();

            A.Outline outline6 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill10 = new A.SolidFill();

            A.SchemeColor schemeColor19 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation12 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset4 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor19.Append(luminanceModulation12);
            schemeColor19.Append(luminanceOffset4);

            solidFill10.Append(schemeColor19);
            A.Round round3 = new A.Round();

            outline6.Append(solidFill10);
            outline6.Append(round3);
            A.EffectList effectList6 = new A.EffectList();

            chartShapeProperties3.Append(outline6);
            chartShapeProperties3.Append(effectList6);

            majorGridlines2.Append(chartShapeProperties3);
            C.NumberingFormat numberingFormat2 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
            C.MajorTickMark majorTickMark2 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
            C.MinorTickMark minorTickMark2 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
            C.TickLabelPosition tickLabelPosition2 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.Low };

            C.ChartShapeProperties chartShapeProperties4 = new C.ChartShapeProperties();
            A.NoFill noFill2 = new A.NoFill();

            A.Outline outline7 = new A.Outline();
            A.NoFill noFill3 = new A.NoFill();

            outline7.Append(noFill3);
            A.EffectList effectList7 = new A.EffectList();

            chartShapeProperties4.Append(noFill2);
            chartShapeProperties4.Append(outline7);
            chartShapeProperties4.Append(effectList7);

            C.TextProperties textProperties2 = new C.TextProperties();
            A.BodyProperties bodyProperties2 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill11 = new A.SolidFill();

            A.SchemeColor schemeColor20 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation13 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset5 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor20.Append(luminanceModulation13);
            schemeColor20.Append(luminanceOffset5);

            solidFill11.Append(schemeColor20);
            A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont4 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties2.Append(solidFill11);
            defaultRunProperties2.Append(latinFont4);
            defaultRunProperties2.Append(eastAsianFont4);
            defaultRunProperties2.Append(complexScriptFont4);

            paragraphProperties2.Append(defaultRunProperties2);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(endParagraphRunProperties2);

            textProperties2.Append(bodyProperties2);
            textProperties2.Append(listStyle2);
            textProperties2.Append(paragraph2);
            C.CrossingAxis crossingAxis2 = new C.CrossingAxis() { Val = (UInt32Value)428045344U };
            C.Crosses crosses2 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
            C.CrossBetween crossBetween2 = new C.CrossBetween() { Val = C.CrossBetweenValues.MidpointCategory };
            C.MajorUnit majorUnit2 = new C.MajorUnit() { Val = 1D };
            C.MinorUnit minorUnit2 = new C.MinorUnit() { Val = 1D };

            valueAxis2.Append(axisId4);
            valueAxis2.Append(scaling2);
            valueAxis2.Append(delete2);
            valueAxis2.Append(axisPosition2);
            valueAxis2.Append(majorGridlines2);
            valueAxis2.Append(numberingFormat2);
            valueAxis2.Append(majorTickMark2);
            valueAxis2.Append(minorTickMark2);
            valueAxis2.Append(tickLabelPosition2);
            valueAxis2.Append(chartShapeProperties4);
            valueAxis2.Append(textProperties2);
            valueAxis2.Append(crossingAxis2);
            valueAxis2.Append(crosses2);
            valueAxis2.Append(crossBetween2);
            valueAxis2.Append(majorUnit2);
            valueAxis2.Append(minorUnit2);

            C.ShapeProperties shapeProperties2 = new C.ShapeProperties();
            A.NoFill noFill4 = new A.NoFill();

            A.Outline outline8 = new A.Outline();
            A.NoFill noFill5 = new A.NoFill();

            outline8.Append(noFill5);
            A.EffectList effectList8 = new A.EffectList();

            shapeProperties2.Append(noFill4);
            shapeProperties2.Append(outline8);
            shapeProperties2.Append(effectList8);

            plotArea1.Append(layout1);
            plotArea1.Append(scatterChart1);
            plotArea1.Append(valueAxis1);
            plotArea1.Append(valueAxis2);
            plotArea1.Append(shapeProperties2);
            C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };
            C.DisplayBlanksAs displayBlanksAs1 = new C.DisplayBlanksAs() { Val = C.DisplayBlanksAsValues.Zero };
            C.ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new C.ShowDataLabelsOverMaximum() { Val = false };

            chart1.Append(autoTitleDeleted1);
            chart1.Append(plotArea1);
            chart1.Append(plotVisibleOnly1);
            chart1.Append(displayBlanksAs1);
            chart1.Append(showDataLabelsOverMaximum1);

            C.ShapeProperties shapeProperties3 = new C.ShapeProperties();

            A.Outline outline9 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill12 = new A.SolidFill();

            A.SchemeColor schemeColor21 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation14 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset6 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor21.Append(luminanceModulation14);
            schemeColor21.Append(luminanceOffset6);

            solidFill12.Append(schemeColor21);
            A.Round round4 = new A.Round();

            outline9.Append(solidFill12);
            outline9.Append(round4);
            A.EffectList effectList9 = new A.EffectList();

            shapeProperties3.Append(outline9);
            shapeProperties3.Append(effectList9);

            C.PrintSettings printSettings1 = new C.PrintSettings();
            C.HeaderFooter headerFooter1 = new C.HeaderFooter();
            C.PageMargins pageMargins2 = new C.PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            C.PageSetup pageSetup2 = new C.PageSetup();

            printSettings1.Append(headerFooter1);
            printSettings1.Append(pageMargins2);
            printSettings1.Append(pageSetup2);

            chartSpace1.Append(date19041);
            chartSpace1.Append(editingLanguage1);
            chartSpace1.Append(roundedCorners1);
            chartSpace1.Append(alternateContent2);
            chartSpace1.Append(chart1);
            chartSpace1.Append(shapeProperties3);
            chartSpace1.Append(printSettings1);

            chartPart1.ChartSpace = chartSpace1;
        }

        // Generates content of imagePart1.
        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
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
            SharedStringTable sharedStringTable1 = new SharedStringTable();

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Microsoft Office User";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2017-03-07T03:50:01Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2017-03-15T19:03:31Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Josue Garcia";
            document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2017-03-07T07:10:05Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        #region Binary Data

        private string spreadsheetPrinterSettingsPart1Data = "UwBuAGEAZwBpAHQAIAAxADMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEEAwbcAAwDQ++ABQEAAQDqCm8IZAABAA8AyAACAAEAyAACAAEATABlAHQAdABlAHIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAgAAAAEAAAD/////AAAAAAAAAAAAAAAAAAAAAERJTlUiALAADAMAAMGVHPsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABgAAAAEAAAAAAAAAAgABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACwAAAAU01USgAAAAAQAKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

    }
}