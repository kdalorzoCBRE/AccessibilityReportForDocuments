using AccessibilityReportForDocuments.core.errors;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;

namespace AccessibilityReportForDocuments.core.scanners
{
    public interface IAccessibilityScanner<T>
    {
        public List<AccessibilityError> Scan(OpenXmlPackage document, T data);
    }

    public abstract class AccessibilityScanner<T>: IAccessibilityScanner<T>
    {
        protected readonly ILogger log;

        public AccessibilityScanner(ILogger log)
        {
            this.log = log;
        }

        public abstract List<AccessibilityError> Scan(OpenXmlPackage document, T data);

    }

}
