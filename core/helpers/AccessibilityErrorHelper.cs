using AccessibilityReportForDocuments.core.errors;
using Microsoft.Extensions.Logging;

namespace AccessibilityReportForDocuments.core.helpers
{
    public static class AccessibilityErrorHelper
    {
        public static void LogAccessibilityError(AccessibilityError error, ILogger log, string className)
        {
            log.LogInformation(className + " found issue on " + error.ObjectName);
        }
    }
}
