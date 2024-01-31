
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

    public class TableHeaderNotFoundError : AccessibilityError
    {
        private readonly string ERROR_DESCRIPTION = "Missing Table Header";
        private readonly string ERROR_TYPE = "Table Header Not Found";

        public TableHeaderNotFoundError(string objectName)
        {
            ObjectName = objectName;
            ErrorType = ERROR_TYPE;
            ErrorDescription = ERROR_DESCRIPTION;
        }
    }

    public class SectionNameNotValidError : AccessibilityError
    {
        private readonly string ERROR_DESCRIPTION = "Default Section Name";
        private readonly string ERROR_TYPE = "Section Name Not Valid";

        public SectionNameNotValidError(string objectName)
        {
            ObjectName = objectName;
            ErrorType = ERROR_TYPE;
            ErrorDescription = ERROR_DESCRIPTION;
        }
    }

    public class SlideTitleNotFound : AccessibilityError
    {
        private readonly string ERROR_DESCRIPTION = "Missing Slide Title";
        private readonly string ERROR_TYPE = "Slide Title Not Found";

        public SlideTitleNotFound(string objectName)
        {
            ObjectName = objectName;
            ErrorType = ERROR_TYPE;
            ErrorDescription = ERROR_DESCRIPTION;
        }
    }

    public class DocumentCorrupted : AccessibilityError
    {
        private readonly string ERROR_DESCRIPTION = "Document is corrupted. Can not be analyzed. This could be because 'access content programatically' has been disabled.";
        private readonly string ERROR_TYPE = "Document Corrupted";

        public DocumentCorrupted(string objectName)
        {
            ObjectName = objectName;
            ErrorType = ERROR_TYPE;
            ErrorDescription = ERROR_DESCRIPTION;
        }
    }
}