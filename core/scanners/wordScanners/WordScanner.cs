using AccessibilityReportForDocuments.core.errors;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace AccessibilityReportForDocuments.core.scanners.wordScanners
{

    public interface IAccessibilityWordScanner<T> : IAccessibilityScanner<T> where T : Body
    {
        public List<AccessibilityError> Scan(OpenXmlPackage document, Body data);
    }
}

