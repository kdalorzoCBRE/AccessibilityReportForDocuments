using AccessibilityReportForDocuments.core.errors;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace AccessibilityReportForDocuments.core.scanners.spreadsheetScanner
{

    internal interface IAccessibilitySpreadSheetScanner<T> : IAccessibilityScanner<T> where T : Workbook
    {
        public List<AccessibilityError> Scan(OpenXmlPackage document, Workbook data);
    }

    internal class SpreadSheetImageAltTextScanner : IAccessibilitySpreadSheetScanner<Workbook>
    {
        public List<AccessibilityError> Scan(OpenXmlPackage document, Workbook data)
        {
            SpreadsheetDocument doc = document as SpreadsheetDocument;

            List<AccessibilityError> imageAltTextNotFoundErrors = new List<AccessibilityError>();

            // TODO: pending 

            return imageAltTextNotFoundErrors;
        }
    }
}

