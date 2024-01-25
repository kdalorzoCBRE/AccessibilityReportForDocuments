using AccessibilityReportForDocuments.core.errors;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Linq;
using Anchor = DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace AccessibilityReportForDocuments.core.scanners.wordScanners
{

    internal static class WordObjectAltTextScanner
    {
        public static List<IAccessibilityWordScanner<Body>> AltTextScanners(ILogger log)
        {
            return new()
            {
                new WordInlineAltTextScanner(log),
                new WordAnchorAltTextScanner(log)
            };
        }
    }

    /// <summary>
    /// Checks Alt Text exists for objects of type Picture, Grahic, Diagram, Chart, Screenshoot, Icon and 3D Models that are inline
    /// </summary>
    internal class WordInlineAltTextScanner : IAccessibilityWordScanner<Body>
    {
        private readonly ILogger log;

        public WordInlineAltTextScanner(ILogger log)
        {
            this.log = log;
        }

        public List<AccessibilityError> Scan(OpenXmlPackage document, Body data)
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
    internal class WordAnchorAltTextScanner : IAccessibilityWordScanner<Body>
    {
        private readonly ILogger log;

        public WordAnchorAltTextScanner(ILogger log)
        {
            this.log = log;
        }

        public List<AccessibilityError> Scan(OpenXmlPackage document, Body data)
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

