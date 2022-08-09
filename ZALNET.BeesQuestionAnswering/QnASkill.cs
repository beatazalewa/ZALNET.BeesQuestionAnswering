using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using System.Collections.Generic;

namespace AzFuncClientsToSPList
{
    public static class QnASkill
    {
        [FunctionName("QnASkill")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string name = req.Query["name"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            name = name ?? data?.name;

            ConfigurationsValues configValues = readEnviornmentVariable();
            GraphServiceClient _graphServiceClient = getGraphClient(configValues);
            var queryOptions = new List<QueryOption>()
                            {
                                new QueryOption("expand", "fields(select=Question,Answer)")
                            };
            IListItemsCollectionPage _listItems;
            try
            {

                _listItems = await _graphServiceClient.Sites[configValues.SiteId].Lists[configValues.clientsListID].Items
                                 .Request(queryOptions)
                                 .GetAsync();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            foreach (var item in _listItems.CurrentPage)
            {
                Console.WriteLine("Question---" + item.Fields.AdditionalData["Question"]);
                Console.WriteLine("Answer---" + item.Fields.AdditionalData["Answer"]);

            }
            return new OkObjectResult("This HTTP triggered function executed successfully");
        }

        public static ConfigurationsValues readEnviornmentVariable()
        {
            ConfigurationsValues configValues = new ConfigurationsValues();

            configValues.Tenantid = System.Environment.GetEnvironmentVariable("tenantid", EnvironmentVariableTarget.Process);
            configValues.Clientid = System.Environment.GetEnvironmentVariable("clientid", EnvironmentVariableTarget.Process);
            configValues.ClientSecret = System.Environment.GetEnvironmentVariable("clientsecret", EnvironmentVariableTarget.Process);
            configValues.SiteId = System.Environment.GetEnvironmentVariable("siteId", EnvironmentVariableTarget.Process);
            configValues.clientsListID = System.Environment.GetEnvironmentVariable("clientsListID", EnvironmentVariableTarget.Process);

            return configValues;
        }

        public static GraphServiceClient getGraphClient(ConfigurationsValues configValues)
        {
            var scopes = new[] { "User.Read" };
            var pca = PublicClientApplicationBuilder
    .Create(configValues.Clientid)
    .WithTenantId(configValues.Tenantid)
    .Build();

            // DelegateAuthenticationProvider is a simple auth provider implementation
            // that allows you to define an async function to retrieve a token
            // Alternatively, you can create a class that implements IAuthenticationProvider
            // for more complex scenarios
            var authProvider = new DelegateAuthenticationProvider(async (request) => {
                // Use Microsoft.Identity.Client to retrieve token
                var result = await pca.AcquireTokenByIntegratedWindowsAuth(scopes).ExecuteAsync();

                request.Headers.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", result.AccessToken);
            });

            GraphServiceClient graphClient = new GraphServiceClient(authProvider);
            return graphClient;
        }
    }

    public class ConfigurationsValues
    {
        public string Clientid { get; set; }
        public string Tenantid { get; set; }
        public string ClientSecret { get; set; }
        public string SiteId { get; set; }
        public string TargetedListId { get; set; }
        public string clientsListID { get; set; }
    }
}