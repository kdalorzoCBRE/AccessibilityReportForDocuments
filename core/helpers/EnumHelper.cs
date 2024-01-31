using AccessibilityReportForDocuments.core.scanners.wordScanners;
using System.ComponentModel;

namespace AccessibilityReportForDocuments.core.helpers
{
    public static class EnumHelper
    {
        public static string Description(this InvalidSectionName val)
        {
            DescriptionAttribute[] attributes = (DescriptionAttribute[])val
               .GetType()
               .GetField(val.ToString())
               .GetCustomAttributes(typeof(DescriptionAttribute), false);
            return attributes.Length > 0 ? attributes[0].Description : string.Empty;
        }

    }
}
