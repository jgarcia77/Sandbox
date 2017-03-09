using DocumentFormat.OpenXml.Packaging;

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
