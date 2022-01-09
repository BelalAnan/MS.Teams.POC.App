using Azure.Messaging.EventGrid;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.EventGrid;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Identity.Client;


namespace MS.Teams.POC.WebApp
{
    public static class SendToTeamsActivity
    {
       // [FunctionName("SendToTeamsActivity")]
        // public static async Task Run([EventGridTrigger] EventGridEvent eventGridEvent, ILogger log,string weburl)
        // {

        //  //   JObject dataObject = eventGridEvent.Data as JObject;
        //    // ActivityDetails details = dataObject.ToObject<ActivityDetails>();

        //     log.LogInformation(eventGridEvent.Data.ToString());
        //     string clientId = Environment.GetEnvironmentVariable("ClientId");
        //     string clientSecret = Environment.GetEnvironmentVariable("ClientSecret");
        //     string tenantId = Environment.GetEnvironmentVariable("TenantId");

        //     string authority = $"https://login.microsoftonline.com/{tenantId}/v2.0";
        //     string userId = details.userId;
        //     string taskId = details.taskId;
        //     string notificationUrl = details.notificationUrl;

        //     var topic = new TeamworkActivityTopic
        //     {        
        //         Source = TeamworkActivityTopicSource.Text,
        //         WebUrl=weburl,
        //         Value = "New Task Created"




        //     };
        //     var activityType = "reservationUpdated";

        //     var previewText = new ItemBody
        //     {
        //         Content = "belal content "
        //     };

        //     var templateParameters = new List<Microsoft.Graph.KeyValuePair>();
        //     await graphServiceClient.Users[userId].Teamwork
        //.SendActivityNotification(topic, activityType, null, previewText, templateParameters)
        //.Request()
        //.PostAsync();
        // }
    }

}
