using GenerateReportPDF.ReportingExecutionService;
using System.Collections.Generic;
using System.IO;
using System.Net;


namespace GenerateReportPDF
{
    /// <summary>
    /// this class is build on code in this article: [Download SQL Server Reporting Service (SSRS) report as PDF using C# (craftedforeveryone.com)](https://www.craftedforeveryone.com/download-sql-server-reporting-service-ssrs-report-as-pdf-using-c-sharp/)
    /// </summary>
    internal class SSRSHandler
    {
        public void Handle()
        {
            string reportPath = Properties.Settings.Default.ReportPath;
            string credentialUsername = Properties.Settings.Default.Username;
            string credentialPassword = Properties.Settings.Default.Password;
            string domain = Properties.Settings.Default.Domain;
            string parameterNameValue = Properties.Settings.Default.ParameterNameValue;
            string pdfFilePath = Properties.Settings.Default.PdfFilePath;

            ReportExecutionService rs = new ReportExecutionService
            {
                Credentials = new NetworkCredential(credentialUsername, credentialPassword, domain),

                ExecutionHeaderValue = new ExecutionHeader()
            };
            _ = new ExecutionInfo();
            _ = rs.LoadReport(reportPath, null);

            List<ParameterValue> parameters = new List<ParameterValue>
            {
                new ParameterValue { Name = "Name", Value = parameterNameValue }
            };

            rs.SetExecutionParameters(parameters.ToArray(), "en-US");

            string deviceInfo = "<DeviceInfo><Toolbar>False</Toolbar></DeviceInfo>";
            var result = rs.Render("PDF", deviceInfo, out _, out _, out _, out _, out _);

            File.WriteAllBytes(pdfFilePath, result);
        }
    }
}
