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
using Microsoft.Extensions.Configuration;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;

namespace appsvc_fnc_dev_bulkuserimport
{
    public static class CheckList
    {
        [FunctionName("CheckList")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            IConfiguration config = new ConfigurationBuilder()

             .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
             .AddEnvironmentVariables()
             .Build();

            var BulkSiteId = config["BulkSiteId"];
            string name = req.Query["name"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            name = name ?? data?.name;

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            string listID = await checkListExist(graphAPIAuth, name, BulkSiteId, log);

            if(listID == "")
            {
                return new BadRequestObjectResult("List do not exist");
            }

            var connectionString = config["AzureWebJobsStorage"];

            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
            CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
            CloudQueue queue = queueClient.GetQueueReference("bulkimportuserlist");
            string ResponsQueue = "";
            ResponsQueue = CreateQueue(queue, listID, BulkSiteId, log).GetAwaiter().GetResult();


            if (String.Equals(ResponsQueue, "Queue create"))
            {
                log.LogInformation("Response queue");
                return new OkObjectResult(ResponsQueue);
            }
            else
            {
                log.LogInformation("Response queue error");

                return new BadRequestObjectResult(ResponsQueue);
            }
        }


        static async Task<string> checkListExist(GraphServiceClient graphAPIAuth, string name, string BulkSiteId, ILogger log)
        {
            string ID = "";
         
            var lists = await graphAPIAuth.Sites[BulkSiteId].Lists
                       .Request()
                       .GetAsync();
            foreach (var item in lists)
            {
                log.LogInformation(item.Name);
                if (item.Name == name){
                    ID = item.Id;
                    break;
                }
            }
            return ID;
        }

        static async Task<string> CreateQueue(CloudQueue theQueue, string listID, string siteID, ILogger log)
        {
            string response = "";
            BulkInfo bulk = new BulkInfo();

            bulk.listID = listID;
            bulk.siteID = siteID;

            string serializedMessage = JsonConvert.SerializeObject(bulk);
            if (await theQueue.CreateIfNotExistsAsync())
            {
                log.LogInformation("The queue was created.");
            }

            CloudQueueMessage message = new CloudQueueMessage(serializedMessage);
            try
            {
                log.LogInformation("create queue");

                await theQueue.AddMessageAsync(message);
                response = "Queue create";
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error in the queue {ex}");
                response = "Queue error";
            }
            return response;
        }
    }
}