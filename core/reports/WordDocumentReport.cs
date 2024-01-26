﻿using AccessibilityReportForDocuments.core.errors;
using AccessibilityReportForDocuments.core.scanners;
using AccessibilityReportForDocuments.core.scanners.wordScanners;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AccessibilityReportForDocuments.core.reports
{
    internal class WordDocumentReport
    {
        private readonly ILogger log;

        private readonly List<IAccessibilityWordScanner<Body>> scanners = new();

        public WordDocumentReport(ILogger log)
        {
            this.log = log;
            scanners.AddRange(WordObjectAltTextScanner.AltTextScanners(this.log));
        }

        public List<AccessibilityError> GenerateReport(Stream stream)
        {
            List<AccessibilityError> accessibilityErrors = new();

            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, false);

            // TODO: validate for null
            Body body = wordDocument.MainDocumentPart.Document.Body;

            foreach (IAccessibilityScanner<Body> scanner in scanners)
            {
                List<AccessibilityError> scannerErrors = scanner.Scan(wordDocument, body);
                accessibilityErrors.AddRange(scannerErrors);
            }

            return accessibilityErrors.GroupBy(x => x.ObjectName).Select(x => x.First()).ToList();
        }
    }
}


