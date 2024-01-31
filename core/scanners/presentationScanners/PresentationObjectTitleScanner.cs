using AccessibilityReportForDocuments.core.errors;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AccessibilityReportForDocuments.core.scanners.wordScanners
{
    public static class PresentationObjectTitleScanner
    {
        public static List<AccessibilityScanner<Presentation>> ObjetTitleScanners(ILogger log)
        {
            return new()
            {
                new SlideTitleScanner(log)
            };
        }
    }

    /// <summary>
    /// Checks for Slide Title 
    /// </summary>
    public class SlideTitleScanner : AccessibilityScanner<Presentation>
    {
        private readonly String TITLE_PLACEHOLDER = "title";

        public SlideTitleScanner(ILogger log) : base(log)
        {
        }

        public override List<AccessibilityError> Scan(OpenXmlPackage document, Presentation data)
        {
            List<AccessibilityError> slideTitleNotFoundErrors = new();

            PresentationDocument doc = document as PresentationDocument;

            int slideIdNumber = 1;

            foreach (SlideId slideId in data.SlideIdList.Cast<SlideId>())
            {
                Slide slide = (doc.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart).Slide;

                var titleShape = slide.Descendants<Shape>()
                    .Where(x => x.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.PlaceholderShape.Type == TITLE_PLACEHOLDER)
                    .FirstOrDefault();


                if (titleShape == null)
                {
                    log.LogInformation(this.GetType().Name + " found issue on Slide" + slideIdNumber);
                    slideTitleNotFoundErrors.Add(new SlideTitleNotFound("Slide " + slideIdNumber));
                }
                else
                {
                    TextBody p = titleShape.Descendants<TextBody>().FirstOrDefault();
                    if (string.IsNullOrEmpty(p.InnerText))
                    {
                        log.LogInformation(this.GetType().Name + " found issue on Slide" + slideIdNumber);
                        slideTitleNotFoundErrors.Add(new SlideTitleNotFound("Slide " + slideIdNumber));

                    }
                }
                slideIdNumber++;
            }


            return slideTitleNotFoundErrors;
        }
    }
}
