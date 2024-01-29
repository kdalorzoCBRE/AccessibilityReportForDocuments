using AccessibilityReportForDocuments.core.errors;
using AccessibilityReportForDocuments.core.scanners.presentationScanners;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Linq;
using GraphicFrame = DocumentFormat.OpenXml.Presentation.GraphicFrame;

namespace AccessibilityReportForDocuments.core.scanners.wordScanners
{

    public static class PresentationObjectHeaderScanner
    {
        public static List<IAccessibilityPresentationScanner<Presentation>> HeaderScanners(ILogger log)
        {
            return new()
            {
                new PresentationTableHeaderScanner(log)
            };
        }
    }

    /// <summary>
    /// Checks header row exists for tables in the presentation. 
    /// </summary>
    public class PresentationTableHeaderScanner : IAccessibilityPresentationScanner<Presentation>
    {
        private readonly ILogger log;

        public PresentationTableHeaderScanner(ILogger log)
        {
            this.log = log;
        }

        public List<AccessibilityError> Scan(OpenXmlPackage document, Presentation data)
        {
            List<AccessibilityError> tableHeaderNotFoundErrors = new();

            PresentationDocument doc = document as PresentationDocument;

            foreach (SlideId slideId in data.SlideIdList.Cast<SlideId>())
            {
                Slide slide = (doc.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart).Slide;

                foreach (var graphicName in slide.Descendants<GraphicFrame>())
                {
                    var graphic = graphicName.Descendants<Graphic>().FirstOrDefault();

                    if (graphic != null)
                    {
                        Table table = graphic.GraphicData.Descendants<Table>().FirstOrDefault();

                        if (table != null)
                        {
                            string tableName = graphicName.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name;

                            if (table.TableProperties.FirstRow == null)
                            {
                                log.LogInformation(this.GetType().Name + " found issue on " + tableName);
                                tableHeaderNotFoundErrors.Add(new TableHeaderNotFoundError(tableName));
                            }
                        }
                    }
                }

            }


            return tableHeaderNotFoundErrors;
        }
    }
}
