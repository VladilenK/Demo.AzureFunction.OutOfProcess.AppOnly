using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using PnP.Core.Admin.Model.SharePoint;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using System.Collections.Specialized;
using System.IO;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using System.Web;
using System;

namespace ProvisioningDemo{
    public class GetSite
    {
        private readonly ILogger logger;
        private readonly IPnPContextFactory contextFactory;
        private readonly AzureFunctionSettings azureFunctionSettings;
        public GetSite(IPnPContextFactory pnpContextFactory, ILoggerFactory loggerFactory, AzureFunctionSettings settings)
        {
            logger = loggerFactory.CreateLogger<CreateSite>();
            contextFactory = pnpContextFactory;
            azureFunctionSettings = settings;
        }

        /// <summary>
        /// Demo function that gets a site collection 
        /// GET/POST url: http://localhost:7071/api/CreateSite?url=yoursiteurl
        /// </summary>
        /// <param name="req"></param>
        /// <returns></returns>
        [Function("GetSite")]
        public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
        {
            logger.LogInformation("GetSite function starting...");

            // Parse the url parameters
            NameValueCollection parameters = HttpUtility.ParseQueryString(req.Url.Query);
            var url = parameters["url"];
            var uri = new Uri(url);

            HttpResponseData response = null;
            logger.LogInformation($"Getting site: {url}");

            try
            {
                using (var pnpContext = await contextFactory.CreateAsync("Default"))
                {
                    response = req.CreateResponse(HttpStatusCode.OK);
                    response.Headers.Add("Content-Type", "application/json");
                    logger.LogInformation($"Getting site details: {url}");
                    //await pnpContext.Site.LoadAsync(p => p.RootWeb.All);
                    //var site = await pnpContext.Site.GetAsync(p => p.All);

                    var siteToCheckDetails = await pnpContext.GetSiteCollectionManager().GetSiteCollectionWithDetailsAsync(uri);

                    await response.WriteStringAsync(JsonSerializer.Serialize(new { 
                        siteId      = siteToCheckDetails.Id,
                        //siteName    = site.RootWeb.Title,
                        //template    = site.RootWeb.WebTemplate,
                        url         = siteToCheckDetails.Url,
                        templateId     = siteToCheckDetails.TemplateId,
                        owner = siteToCheckDetails.SiteOwnerName,
                        createdBy = siteToCheckDetails.CreatedBy
                    }));


                    return response;
                }
            }
            catch (Exception ex)
            {
                response = req.CreateResponse(HttpStatusCode.OK);
                response.Headers.Add("Content-Type", "application/json");
                await response.WriteStringAsync(JsonSerializer.Serialize(new { error = ex.Message }));
                return response;
            }
        }
    }
}
