using AccessibilityReportForDocuments.core.errors;
using AccessibilityReportForDocuments.core.helpers;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;

namespace AccessibilityReportForDocuments.core.scanners.wordScanners
{

    public enum InvalidSectionName
    {
        [Description("Default Section")]
        DefaultSection,
        [Description("Untitled Section")]
        UntitledSection,
        [Description("Section #")]
        SectionNumber

    }
    public static class PresentationSectionNameScanner
    {
        public static List<AccessibilityScanner<Presentation>> SectionNameScanners(ILogger log)
        {
            return new()
            {
                new SectionNameScanner(log)
            };
        }
    }

    /// <summary>
    /// Checks header row exists for tables in the presentation. 
    /// </summary>
    public class SectionNameScanner : AccessibilityScanner<Presentation>
    {
        public SectionNameScanner(ILogger log) : base(log)
        {
        }

        public override List<AccessibilityError> Scan(OpenXmlPackage document, Presentation data)
        {
            List<AccessibilityError> sectionNameNotFound = new();

            SectionList sectionList = data.Descendants<SectionList>().FirstOrDefault();

            foreach (var x in sectionList.Descendants<Section>())
            {
                string name = x.Name;
                //Search on predefined invalid names 
                foreach (InvalidSectionName y in Enum.GetValues(typeof(InvalidSectionName)))
                {
                    if (y.Description() == name)
                    {
                        log.LogInformation(this.GetType().Name + " found issue on " + name);
                        sectionNameNotFound.Add(new SectionNameNotValidError(name));
                    }
                }                
                // Search for regex Section # 
                Regex regex = new Regex("Section [0-9]+");
                if (regex.Match(name).Success)
                {
                    sectionNameNotFound.Add(new SectionNameNotValidError(name));
                }
            }
            return sectionNameNotFound;
        }
    }
}
