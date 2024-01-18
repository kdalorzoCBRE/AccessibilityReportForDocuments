using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Collections.Generic;
using AccessibilityReportForDocuments.core.errors;
using AccessibilityReportForDocuments.core.reports;


namespace AccessibilityReportForDocuments.api
{

    public static class WordDocumentAccessibilityReport
    {

        [FunctionName("WordDocumentAccessibilityReport")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);

            log.LogInformation("Processing request content.");
            string documentContent = data?.body?["$content"];
            byte[] documentContentBytes = Convert.FromBase64String(documentContent);
            Stream stream = new MemoryStream(documentContentBytes);

            log.LogInformation("Generating report for Word Document.");
            WordDocumentReport report = new(log);
            List<AccessibilityError> result = report.GenerateReport(stream);
            log.LogInformation($"Report for Word Document generated: {result.Count} accessibility errors found.");

            string responseMessage = JsonConvert.SerializeObject(result);
            return new OkObjectResult(responseMessage);
        }
    }
}
