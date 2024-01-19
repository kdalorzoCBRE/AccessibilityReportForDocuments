using AccessibilityReportForDocuments.core.errors;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;
using Anchor = DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace AccessibilityReportForDocuments.core.scanners.wordScanners
{

    internal interface IAccessibilityWordScanner<T> : IAccessibilityScanner<T> where T : Body
    {
        public List<AccessibilityError> Scan(OpenXmlPackage document, Body data);
    }

    /// <summary>
    /// Checks for Alt Text for objects of type Picture, Grahic, Diagram, Chart and Screenshoot
    /// </summary>
    internal class WordInlineAltTextScanner : IAccessibilityWordScanner<Body>
    {
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
                            inlineAltTextNotFoundErrors.Add(new ObjectAltTextNotFoundError(name));
                        }
                    }
                }
            }
            return inlineAltTextNotFoundErrors;
        }
    }

    /// <summary>
    /// Checks for Alt Text for objects of type Icon and 3D Model
    /// </summary>
    internal class WordAnchorAltTextScanner : IAccessibilityWordScanner<Body>
    {
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
                            anchorAltTextNotFoundErrors.Add(new ObjectAltTextNotFoundError(name));
                        }

                    }
                }
            }
            return anchorAltTextNotFoundErrors;
        }
    }
}

