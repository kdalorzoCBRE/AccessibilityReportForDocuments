using AccessibilityReportForDocuments.core.errors;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;
using Picture = DocumentFormat.OpenXml.Drawing.Pictures.Picture;

namespace AccessibilityReportForDocuments.core.scanners.wordScanners
{

    internal interface IAccessibilityWordScanner<T> : IAccessibilityScanner<T> where T : Body
    {
        public List<AccessibilityError> Scan(OpenXmlPackage document, Body data);
    }

    internal class WordImageAltTextScanner : IAccessibilityWordScanner<Body>
    {
        public List<AccessibilityError> Scan(OpenXmlPackage document, Body data)
        {
            List<AccessibilityError> imageAltTextNotFoundErrors = new();

            foreach (var x in data.Descendants<Paragraph>())
            {
                foreach (Run run in x.Descendants<Run>())
                {
                    Picture image = run.Descendants<Picture>().FirstOrDefault();

                    if (image != null)
                    {
                        var name = image.NonVisualPictureProperties.NonVisualDrawingProperties.Name;
                        var altText = image.NonVisualPictureProperties.NonVisualDrawingProperties.Description;

                        if (altText == null)
                        {
                            imageAltTextNotFoundErrors.Add(new ImageAltTextNotFoundError(name));
                        }
                    }
                }
            }
            return imageAltTextNotFoundErrors;
        }
    }
}

