using AccessibilityReportForDocuments.core.errors;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Linq;


namespace AccessibilityReportForDocuments.core.scanners.wordScanners
{

    internal static class WordObjectAltTextScanner
    {
        public static List<AccessibilityScanner<Body>> AltTextScanners(ILogger log)
        {
            return new()
            {
                new InlineAltTextScanner(log),
                new WordAnchorAltTextScanner(log)
            };
        }
    }

    /// <summary>
    /// Checks Alt Text exists for objects of type Picture, Grahic, Diagram, Chart, Screenshoot, Icon and 3D Models that are inline
    /// </summary>
    internal class InlineAltTextScanner : AccessibilityScanner<Body>
    {
        public InlineAltTextScanner(ILogger log) : base(log)
        {
        }

        public override List<AccessibilityError> Scan(OpenXmlPackage document, Body data)
        {
            List<AccessibilityError> inlineAltTextNotFoundErrors = new();

            foreach (Paragraph paragraph in data.Descendants<Paragraph>())
            {
                foreach (Run run in paragraph.Descendants<Run>())
                {
                    Inline inline = run.Descendants<Inline>().FirstOrDefault();

                    if (inline != null)
                    {
                        string altText = inline.DocProperties.Description;
                        string name = inline.DocProperties.Name;

                        if (altText == null)
                        {
                            log.LogInformation(this.GetType().Name + " found issue on " + name);
                            inlineAltTextNotFoundErrors.Add(new ObjectAltTextNotFoundError(name));
                        }
                    }
                }
            }
            return inlineAltTextNotFoundErrors;
        }
    }

    /// <summary>
    /// Checks Alt Text exists for objects of type Picture, Grahic, Diagram, Chart, Screenshoot, Icon and 3D Models that are not inline
    /// </summary>
    internal class WordAnchorAltTextScanner : AccessibilityScanner<Body>
    {
        public WordAnchorAltTextScanner(ILogger log) : base(log)
        {
        }

        public override List<AccessibilityError> Scan(OpenXmlPackage document, Body data)
        {
            List<AccessibilityError> anchorAltTextNotFoundErrors = new();

            foreach (Paragraph paragraph in data.Descendants<Paragraph>())
            {
                foreach (Run run in paragraph.Descendants<Run>())
                {
                    Anchor anchor = run.Descendants<Anchor>().FirstOrDefault();

                    if (anchor != null)
                    {
                        var docProperties = anchor.GetFirstChild<DocProperties>();

                        string altText = docProperties.Description;
                        string name = docProperties.Name;

                        if (altText == null)
                        {
                            log.LogInformation(this.GetType().Name + " found issue on " + name);
                            anchorAltTextNotFoundErrors.Add(new ObjectAltTextNotFoundError(name));
                        }

                    }
                }
            }
            return anchorAltTextNotFoundErrors;
        }
    }
}

