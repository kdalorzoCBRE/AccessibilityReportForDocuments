using System.Linq;


namespace AccessibilityReportForDocuments.core.errors
{
    public class AccessibilityError
    {
        public string ObjectName { get; set; }
        public string ErrorType { get; set; }
        public string ErrorDescription { get; set; }

    }

    public class ObjectAltTextNotFoundError : AccessibilityError
    {
        private readonly string ERROR_DESCRIPTION = "Missing Object Description";
        private readonly string ERROR_TYPE = "Object Alt Text Not Found";

        public ObjectAltTextNotFoundError(string objectName)
        {
            ObjectName = objectName;
            ErrorType = ERROR_TYPE;
            ErrorDescription = ERROR_DESCRIPTION;
        }
    }
}