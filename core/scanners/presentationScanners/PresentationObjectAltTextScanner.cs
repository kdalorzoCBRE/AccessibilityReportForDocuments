﻿using AccessibilityReportForDocuments.core.errors;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Linq;

namespace AccessibilityReportForDocuments.core.scanners.presentationScanners
{
    public static class PresentationObjectAltTextScanner
    {
        public static List<AccessibilityScanner<Presentation>> AltTextScanners(ILogger log)
        {
            return new()
            {
                new ImageAltTextScanner(log),
                new ShapeAltTextScanner(log),
                new GraphicAltTextScanner(log)
            };
        }
    }
    /// <summary>
    /// Checks Alt Text exists for objects of type Picture, Screenshot, Icon and 3D Models
    /// </summary>
    public class ImageAltTextScanner : AccessibilityScanner<Presentation>
    {
        public ImageAltTextScanner(ILogger log) : base(log)
        {
        }

        public override List<AccessibilityError> Scan(OpenXmlPackage document, Presentation data)
        {
            PresentationDocument doc = document as PresentationDocument;

            List<AccessibilityError> imageAltTextNotFoundErrors = new();

            foreach (SlideId slideId in data.SlideIdList.Cast<SlideId>())
            {
                Slide slide = (doc.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart).Slide;

                foreach (var image in slide.Descendants<Picture>())
                {
                    var name = image.NonVisualPictureProperties.NonVisualDrawingProperties.Name;
                    var altText = image.NonVisualPictureProperties.NonVisualDrawingProperties.Description;

                    if (altText == null)
                    {
                        log.LogInformation(this.GetType().Name + " found issue on " + name);
                        imageAltTextNotFoundErrors.Add(new ObjectAltTextNotFoundError(name));
                    }
                }
            }
            return imageAltTextNotFoundErrors;
        }
    }

    /// <summary>
    /// Checks Alt Text exists for objects of type Figure 
    /// </summary>
    public class ShapeAltTextScanner : AccessibilityScanner<Presentation>
    {
        public ShapeAltTextScanner(ILogger log) : base(log)
        {
        }

        public override List<AccessibilityError> Scan(OpenXmlPackage document, Presentation data)
        {
            PresentationDocument doc = document as PresentationDocument;

            List<AccessibilityError> shapeAltTextNotFoundErrors = new();

            foreach (SlideId slideId in data.SlideIdList.Cast<SlideId>())
            {
                Slide slide = (doc.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart).Slide;

                foreach (var shape in slide.Descendants<Shape>())
                {
                    if (shape.ShapeStyle != null)
                    {
                        var name = shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name;
                        var altText = shape.NonVisualShapeProperties.NonVisualDrawingProperties.Description;

                        if (altText == null)
                        {
                            log.LogInformation(this.GetType().Name + " found issue on " + name);
                            shapeAltTextNotFoundErrors.Add(new ObjectAltTextNotFoundError(name));
                        }
                    }
                }
            }
            return shapeAltTextNotFoundErrors;
        }
    }

    /// <summary>
    /// Checks Alt Text exists for objects of type 3D Model, SmartArt and Chart
    /// </summary>
    public class GraphicAltTextScanner : AccessibilityScanner<Presentation>
    {
        public GraphicAltTextScanner(ILogger log) : base(log)
        {
        }

        public override List<AccessibilityError> Scan(OpenXmlPackage document, Presentation data)
        {
            PresentationDocument doc = document as PresentationDocument;

            List<AccessibilityError> shapeAltTextNotFoundErrors = new();

            foreach (SlideId slideId in data.SlideIdList.Cast<SlideId>())
            {
                Slide slide = (doc.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart).Slide;

                foreach (GraphicFrame shape in slide.Descendants<GraphicFrame>())
                {
                    var name = shape.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name;
                    var altText = shape.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Description;

                    if (altText == null)
                    {
                        log.LogInformation(this.GetType().Name + " found issue on " + name);
                        shapeAltTextNotFoundErrors.Add(new ObjectAltTextNotFoundError(name));
                    }
                }
            }
            return shapeAltTextNotFoundErrors;
        }
    }
}
