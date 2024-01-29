using AccessibilityReportForDocuments.core.errors;
using AccessibilityReportForDocuments.core.scanners.presentationScanners;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;

namespace AccessibilityReportForDocuments.core.scanners.wordScanners
{

    public static class PresentationSectionNameScanner
    {
        public static List<AccessibilityScanner<Presentation>> SectionNameScanners(ILogger log)
        {
            return new()
            {
                new SectionNameScanner(log)
            };
        }
    }

    /// <summary>
    /// Checks header row exists for tables in the presentation. 
    /// </summary>
    public class SectionNameScanner : AccessibilityScanner<Presentation>
    {
        public SectionNameScanner(ILogger log) : base(log)
        {
        }

        public override List<AccessibilityError> Scan(OpenXmlPackage document, Presentation data)
        {
            List<AccessibilityError> sectionNameNotFound = new();

            PresentationDocument doc = document as PresentationDocument;

            return sectionNameNotFound;
        }
    }
}
