using AccessibilityReportForDocuments.core.errors;
using AccessibilityReportForDocuments.core.scanners;
using AccessibilityReportForDocuments.core.scanners.wordScanners;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AccessibilityReportForDocuments.core.reports
{
    public class WordDocumentReport
    {
        private readonly ILogger log;

        private readonly List<AccessibilityScanner<Body>> scanners = new();

        public WordDocumentReport(ILogger log)
        {
            this.log = log;
            scanners.AddRange(WordObjectAltTextScanner.AltTextScanners(this.log));
            scanners.AddRange(WordObjectHeaderScanner.HeaderScanners(this.log));
        }

        public List<AccessibilityError> GenerateReport(Stream stream)
        {
            List<AccessibilityError> accessibilityErrors = new();

            try
            {
                using WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, false);

                // TODO: validate for null
                Body body = wordDocument.MainDocumentPart.Document.Body;

                foreach (IAccessibilityScanner<Body> scanner in scanners)
                {
                    List<AccessibilityError> scannerErrors = scanner.Scan(wordDocument, body);
                    accessibilityErrors.AddRange(scannerErrors);
                }
            }
            catch(FileFormatException ex)
            {
                log.LogInformation("Document Corrupted: " + ex.Message);
                accessibilityErrors.Add(new DocumentCorrupted("Word"));
            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
            }

            return accessibilityErrors.GroupBy(x => x.ObjectName).Select(x => x.First()).ToList();
        }
    }
}


