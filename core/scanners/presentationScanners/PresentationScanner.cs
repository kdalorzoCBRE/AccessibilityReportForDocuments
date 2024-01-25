using AccessibilityReportForDocuments.core.errors;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Collections.Generic;

namespace AccessibilityReportForDocuments.core.scanners.presentationScanners
{

    internal interface IAccessibilityPresentationScanner<T> : IAccessibilityScanner<T> where T : Presentation
    {
        public List<AccessibilityError> Scan(OpenXmlPackage document, Presentation data);
    }
    
}

