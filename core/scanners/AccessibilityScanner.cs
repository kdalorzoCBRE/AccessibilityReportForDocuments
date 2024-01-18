using AccessibilityReportForDocuments.core.errors;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;

namespace AccessibilityReportForDocuments.core.scanners
{
    internal interface IAccessibilityScanner<T>
    {
        public List<AccessibilityError> Scan(OpenXmlPackage document, T data);
    }
}
