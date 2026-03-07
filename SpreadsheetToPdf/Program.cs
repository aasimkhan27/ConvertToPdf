using System;
using System.Configuration;
using System.Web.Http;
using System.Web.Http.SelfHost;
using SpreadsheetToPdf.App_Start;
using System.IO;

namespace SpreadsheetToPdf
{
    internal static class Program
    {
        [STAThread]
        private static int Main()
        {
            string baseAddress = ConfigurationManager.AppSettings["BaseAddress"] ?? "http://localhost:5000";

            var selfHostConfiguration = new HttpSelfHostConfiguration(baseAddress)
            {
                MaxReceivedMessageSize = 50L * 1024L * 1024L
            };

            selfHostConfiguration.TransferMode = System.ServiceModel.TransferMode.Streamed;
            WebApiConfig.Register(selfHostConfiguration);

            using (var server = new HttpSelfHostServer(selfHostConfiguration))
            {
                try
                {
                    server.OpenAsync().Wait();
                    Console.WriteLine("SpreadsheetToPdf Web API started.");
                    Console.WriteLine("Listening on: " + baseAddress);
                    Console.WriteLine("Health endpoint: " + baseAddress.TrimEnd('/') + "/api/conversion/health");
                    Console.WriteLine("Press ENTER to stop.");
                    Console.ReadLine();
                    return 0;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("Failed to start Web API host: " + ex.Message);
                    return 1;
                }
            }
        }
    }
}
