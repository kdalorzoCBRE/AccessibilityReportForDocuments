using AccessibilityReportForDocuments.core.errors;
using AccessibilityReportForDocuments.core.scanners;
using AccessibilityReportForDocuments.core.scanners.wordScanners;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.IO;

namespace AccessibilityReportForDocuments.core.reports
{
    internal class WordDocumentReport
    {
        private readonly ILogger log;

        private readonly List<IAccessibilityWordScanner<Body>> scanners = new()
        {
            new WordInlineAltTextScanner(),
            new WordAnchorAltTextScanner()
        };

        public WordDocumentReport(ILogger log)
        {
            this.log = log;
        }

        public List<AccessibilityError> GenerateReport(Stream stream)
        {
            List<AccessibilityError> accessibilityErrors = new();

            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, false);

            Body body = wordDocument.MainDocumentPart.Document.Body;

            foreach (IAccessibilityScanner<Body> scanner in scanners)
            {
                List<AccessibilityError> scannerErrors = scanner.Scan(wordDocument, body);
                accessibilityErrors.AddRange(scannerErrors);
            }

            return accessibilityErrors;
        }
    }
}


