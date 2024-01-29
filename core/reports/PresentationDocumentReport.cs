using AccessibilityReportForDocuments.core.errors;
using AccessibilityReportForDocuments.core.scanners;
using AccessibilityReportForDocuments.core.scanners.presentationScanners;
using AccessibilityReportForDocuments.core.scanners.wordScanners;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.IO;
using System.Linq;


namespace AccessibilityReportForDocuments.core.reports
{
    internal class PresentationDocumentReport
    {
        private readonly ILogger log;

        private readonly List<AccessibilityScanner<Presentation>> scanners = new();
 

        public PresentationDocumentReport(ILogger log)
        {
            this.log = log;
            scanners.AddRange(PresentationObjectAltTextScanner.AltTextScanners(this.log));
            scanners.AddRange(PresentationObjectHeaderScanner.HeaderScanners(this.log));
        }

        public List<AccessibilityError> GenerateReport(Stream stream)
        {
            List<AccessibilityError> accessibilityErrors = new();

            using PresentationDocument presentationDocument = PresentationDocument.Open(stream, false);

            // TODO: validate for null
            Presentation presentation = presentationDocument.PresentationPart.Presentation;

            foreach (IAccessibilityScanner<Presentation> scanner in scanners)
            {
                List<AccessibilityError> scannerErrors = scanner.Scan(presentationDocument, presentation);
                accessibilityErrors.AddRange(scannerErrors);
            }
            return accessibilityErrors.GroupBy(x => x.ObjectName).Select(x => x.First()).ToList();
        }
    }
}


