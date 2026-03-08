using System.Web;
using System.Web.Http;
using SpreadsheetToPdf.App_Start;

namespace SpreadsheetToPdf
{
    public class WebApiApplication : HttpApplication
    {
        protected void Application_Start()
        {
            GlobalConfiguration.Configure(WebApiConfig.Register);
        }
    }
}
