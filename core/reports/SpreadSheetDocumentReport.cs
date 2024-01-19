using AccessibilityReportForDocuments.core.errors;
using AccessibilityReportForDocuments.core.scanners;
using AccessibilityReportForDocuments.core.scanners.spreadsheetScanner;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.IO;

namespace AccessibilityReportForDocuments.core.reports
{
    internal class SpreadSheetDocumentReport
    {
        private readonly ILogger log;

        private readonly List<IAccessibilitySpreadSheetScanner<Workbook>> scanners = new()
        {
            new SpreadSheetImageAltTextScanner()
        };

        public SpreadSheetDocumentReport(ILogger log)
        {
            this.log = log;
        }

        public List<AccessibilityError> GenerateReport(Stream stream)
        {
            List<AccessibilityError> accessibilityErrors = new();

            using SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(stream, false);

            Workbook workbook = spreadSheetDocument.WorkbookPart.Workbook;

            foreach (IAccessibilityScanner<Workbook> scanner in scanners)
            {
                List<AccessibilityError> scannerErrors = scanner.Scan(spreadSheetDocument, workbook);
                accessibilityErrors.AddRange(scannerErrors);
            }

            return accessibilityErrors;
        }
    }
}


