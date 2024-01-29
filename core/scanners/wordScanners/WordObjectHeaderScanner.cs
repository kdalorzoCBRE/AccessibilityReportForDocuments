using AccessibilityReportForDocuments.core.errors;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Linq;

namespace AccessibilityReportForDocuments.core.scanners.wordScanners
{

    public static class WordObjectHeaderScanner
    {
        public static List<AccessibilityScanner<Body>> HeaderScanners(ILogger log)
        {
            return new()
            {
                new WordTableHeaderScanner(log)         
            };
        }
    }

    /// <summary>
    /// Checks header row eists for tables in the document. 
    /// </summary>
    public class WordTableHeaderScanner : AccessibilityScanner<Body>
    {
        public WordTableHeaderScanner(ILogger log) : base(log)
        {
        }

        public override List<AccessibilityError> Scan(OpenXmlPackage document, Body data)
        {
            List<AccessibilityError> tableHeaderNotFoundErrors = new();

            foreach (Table table in  data.Elements<Table>())
            {
                TableProperties tableProperties = table.Descendants<TableProperties>().FirstOrDefault();                

                TableLook tableLook = tableProperties.TableLook;
                string tableName = "Table " + tableLook.Val;

                if ( tableLook.FirstRow == false)
                {
                    log.LogInformation(this.GetType().Name + " found issue on " + tableName);
                    tableHeaderNotFoundErrors.Add(new TableHeaderNotFoundError(tableName));
                }
            }
            return tableHeaderNotFoundErrors;
        }
    }
}

