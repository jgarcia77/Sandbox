using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;

namespace Sandbox.OpenXML
{
    public class AnalysisWorksheets
    {
        private OverviewWorksheet _overviewWorksheet;
        private DataModelWorksheet _dataModelWorksheet;
        private ResultsReportWorksheet _resultsReportWorksheet;

        public AnalysisWorksheets(int sequence)
        {
            _overviewWorksheet = new OverviewWorksheet(sequence);
            _dataModelWorksheet = new DataModelWorksheet(sequence);
            _resultsReportWorksheet = new ResultsReportWorksheet(sequence);
        }

        public void AppendTo(WorkbookPart workbookPart)
        {
            _resultsReportWorksheet.AppendTo(workbookPart);

            _dataModelWorksheet.AppendTo(workbookPart, _resultsReportWorksheet.ImagePart);

            _overviewWorksheet.AppendTo(workbookPart, _resultsReportWorksheet.ImagePart);
        }
    }
}
