using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;
using SpreadsheetToPdf.Core;

namespace SpreadsheetToPdf.Controllers
{
    [RoutePrefix("api/conversion")]
    public sealed class ConversionController : ApiController
    {
        private readonly SpreadsheetConversionService _conversionService = new SpreadsheetConversionService();

        [HttpGet]
        [Route("health")]
        public IHttpActionResult Health()
        {
            return Ok("SpreadsheetToPdf Web API is running.");
        }

        [HttpPost]
        [Route("pdf")]
        public async Task<HttpResponseMessage> ConvertToPdf()
        {
            if (!Request.Content.IsMimeMultipartContent())
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Request must be multipart/form-data.");
            }

            string tempInputPath = null;

            try
            {
                MultipartMemoryStreamProvider provider = await Request.Content.ReadAsMultipartAsync();

                string worksheetName = null;
                HttpContent fileContent = null;
                string originalFileName = null;

                foreach (HttpContent content in provider.Contents)
                {
                    string contentDispositionName = content.Headers.ContentDisposition?.Name?.Trim('"');
                    if (string.Equals(contentDispositionName, "worksheetName", StringComparison.OrdinalIgnoreCase))
                    {
                        worksheetName = (await content.ReadAsStringAsync()).Trim();
                        worksheetName = string.IsNullOrWhiteSpace(worksheetName) ? null : worksheetName;
                        continue;
                    }

                    if (string.Equals(contentDispositionName, "file", StringComparison.OrdinalIgnoreCase))
                    {
                        fileContent = content;
                        originalFileName = content.Headers.ContentDisposition?.FileName?.Trim('"');
                    }
                }

                if (fileContent == null)
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Missing 'file' form-data part.");
                }

                byte[] fileBytes = await fileContent.ReadAsByteArrayAsync();
                if (fileBytes == null || fileBytes.Length == 0)
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded file is empty.");
                }

                originalFileName = string.IsNullOrWhiteSpace(originalFileName) ? "upload.xlsx" : Path.GetFileName(originalFileName);
                tempInputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + Path.GetExtension(originalFileName));
                File.WriteAllBytes(tempInputPath, fileBytes);

                ConversionResult result = _conversionService.ConvertToPdf(tempInputPath, worksheetName);

                var response = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(result.Content)
                };

                response.Content.Headers.ContentType = new MediaTypeHeaderValue(result.ContentType);
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = result.FileName
                };
                response.Headers.Add("X-Used-Fallback", result.UsedFallback.ToString());
                response.Headers.Add("X-Conversion-Message", result.Message);
                return response;
            }
            catch (WorksheetNotFoundException ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.NotFound, ex.Message);
            }
            catch (InvalidDataException ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
            catch (ExcelNotInstalledException ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.ServiceUnavailable, ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.Forbidden, ex.Message);
            }
            catch (ExcelInteropException ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
            }
            finally
            {
                TryDeleteFile(tempInputPath);
            }
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }
    }
}
