using AccessibilityReportForDocuments.core.errors;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Collections.Generic;
using System.Linq;

namespace AccessibilityReportForDocuments.core.scanners.presentationScanners
{

    internal interface IAccessibilityPresentationScanner<T> : IAccessibilityScanner<T> where T : Presentation
    {
        public List<AccessibilityError> Scan(OpenXmlPackage document, Presentation data);
    }

    internal class PresentationImageAltTextScanner : IAccessibilityPresentationScanner<Presentation>
    {
        public List<AccessibilityError> Scan(OpenXmlPackage document, Presentation data)
        {
            PresentationDocument doc = document as PresentationDocument;

            List<AccessibilityError> imageAltTextNotFoundErrors = new();

            foreach (SlideId slideId in data.SlideIdList.Cast<SlideId>()) 
            {
                Slide slide = (doc.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart).Slide;

                // OR: https://stackoverflow.com/questions/32009006/openxml-get-image-alt-text-title

                foreach (var image in slide.Descendants<Picture>())
                {
                    var name = image.NonVisualPictureProperties.NonVisualDrawingProperties.Name;
                    var altText = image.NonVisualPictureProperties.NonVisualDrawingProperties.Description;

                    if (altText == null)
                    {
                        imageAltTextNotFoundErrors.Add(new ImageAltTextNotFoundError(name));
                    }
                }
            }
            return imageAltTextNotFoundErrors;
        }
    }
}

