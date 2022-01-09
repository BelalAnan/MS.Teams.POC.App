// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace MS.Teams.POC.WebApp
{
    /// <summary>
    /// Helper class to call a protected API and process its result
    /// </summary>
    public class ProtectedApiCallHelper
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="httpClient">HttpClient used to call the protected API</param>
        public ProtectedApiCallHelper(HttpClient httpClient)
        {
            HttpClient = httpClient;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="httpClient">HttpClient used to call the protected API</param>
        public ProtectedApiCallHelper(GraphServiceClient graphServiceClient)
        {
            GraphServiceClient = graphServiceClient;
        }
        protected HttpClient HttpClient { get; private set; }
        protected GraphServiceClient GraphServiceClient { get; private set; }
        public async Task<IGraphServiceApplicationsCollectionPage> GetApplicationAsync()
        {
            return await GraphServiceClient.Applications
             .Request()
             .GetAsync();
        }
 
        public async Task<User> GetUser(string email)
        {
            var users = await GraphServiceClient.Users.Request().Filter($"mail eq '{email}'").GetAsync();
            return users.FirstOrDefault();
        }
        public async Task<IUserChatsCollectionPage> ListAllChats(string userid)
        {
        return await GraphServiceClient.Users[userid].Chats
              .Request()
              .GetAsync();
        }

        //public async Task<ChatMessage> CreateChannelMessage(string content,string channelid,string teamid)
        //{

        //    var chatMessage = new ChatMessage
        //    {

        //        Body = new ItemBody
        //        {
        //            Content = content,

        //        },

        //    };         
        //     await GraphServiceClient.Teams[teamid].Channels[channelid].Messages
        //        .Request()
        //        .AddAsync(chatMessage);
        //    return chatMessage;
        //}
        public async Task<ChatMessage> CreateChannelMessage(string content, string channelid, string teamid,string userid)
        {

            var chatMessage = new ChatMessage
            {

                Body = new ItemBody
                {
                    Content = content,

                },

            };
          var result=  await GraphServiceClient.Teams[teamid].Channels[channelid].Messages
               .Request()
               .AddAsync(chatMessage);
         await   CreateNotification(userid,"belal notifiy from post", teamid, content, result.Id,result.WebUrl);
            return result;

        }
        public async Task<Chat> CreateOneToOneChat(string SendTo)
        {

            var chat = new Chat
            {
                ChatType = ChatType.OneOnOne,
                Members = new ChatMembersCollectionPage()
    {
        new AadUserConversationMember
        {
            Roles = new List<String>()
            {
                "owner"
            },
            AdditionalData = new Dictionary<string, object>()
            {
                {"user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{SendTo}')"}
            }
        },
        new AadUserConversationMember
        {
            Roles = new List<String>()
            {
                "owner"
            },
            AdditionalData = new Dictionary<string, object>()
            {
                {"user@odata.bind", "https://graph.microsoft.com/v1.0/users('7ff19822-3bf2-4c97-a760-a059c60b97c6')"}
            }
        }
    }
            };

          return  await GraphServiceClient.Chats
                .Request()
                .AddAsync(chat);

        }

        public async Task<Calendar> CreateUserCalendar(string Id, string title)
        {

            var calendar = new Calendar
            {
                Name = title
            };

           return await GraphServiceClient.Users[Id].Calendars
                .Request()
                .AddAsync(calendar);

            
        }
        public async Task AddMemberToTeam(string memberid,string teamid)
        {
            var conversationMember = new AadUserConversationMember
            {
                Roles = new List<String>()
    {
        "owner"
    },
                AdditionalData = new Dictionary<string, object>()
    {
        {"user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{memberid}')"}
    }
            };

            await GraphServiceClient.Teams[teamid].Members
                .Request()
                .AddAsync(conversationMember);
        }
        public async Task<IChannelMessagesCollectionPage> GetAllMessagesAtChannel(string teamid, string channelid)
        {

            return await GraphServiceClient.Teams[teamid].Channels[channelid].Messages
                   .Request()
                  .GetAsync();
        }
      
        public  async Task<Team> CreateTeam(string DisplayName, string Description)
        {
            var team = new Team
            {
                DisplayName = DisplayName,
                Description = Description,
                Members = new TeamMembersCollectionPage()
    {
        new AadUserConversationMember
        {
            Roles = new List<String>()
            {
                "owner"
            },
            AdditionalData = new Dictionary<string, object>()
            {
                {"user@odata.bind", "https://graph.microsoft.com/v1.0/users('7ff19822-3bf2-4c97-a760-a059c60b97c6')"}
            }
        }

    },
                AdditionalData = new Dictionary<string, object>()
    {
        {"template@odata.bind", "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"}
    }
            };

           return  await GraphServiceClient.Teams
                .Request()
                .AddAsync(team);

        }
        public async Task<Channel> CreateChannel(string DisplayName, string Description, string id,string userid)
        {
            var channel = new Channel
            {
                MembershipType = ChannelMembershipType.Standard,
                DisplayName = DisplayName,
                Description = Description,
                Members = new ChannelMembersCollectionPage()
    {
        new AadUserConversationMember
        {
            Roles = new List<String>()
            {
                "owner"
            },
            AdditionalData = new Dictionary<string, object>()
            {
                {"user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{userid}')"}
            }
        }
    }
            };
         return  await GraphServiceClient.Teams[id].Channels.Request().AddAsync(channel);

        }
        public async Task<IList<Team>> GetTeams(string Id)
        {
           var res= await GraphServiceClient.Users[Id].JoinedTeams.Request().GetAsync();
            return res;
        }
        public async Task<IUserCalendarsCollectionPage> ListUserCalendrs(string userid)
        {
           return await GraphServiceClient.Users[userid].Calendars.Request().GetAsync();
        }
         public async Task<IList<Channel>> GetChannels(string teamid)
        {
            var channels = await GraphServiceClient.Teams[teamid].Channels.Request().GetAsync();
            return channels;
        }
        public async Task DeleteCalendar(string calendarid, string userid)
        {

         await GraphServiceClient.Users[userid].Calendars[calendarid]
       .Request()
      .DeleteAsync();
 
        }
        public async Task<Event> CreateUserEvent(string Id, string calendarName, Event @event)
        {
             var Calendars = await GraphServiceClient.Users[Id].Calendars.Request().Filter($"name eq '{calendarName}'").GetAsync();



            var calendar = Calendars.FirstOrDefault();           
             var events = await GraphServiceClient.Users[Id].Calendars[calendar.Id].Events.Request().AddAsync(@event);


            return events;
        }
        public async Task<ChatMessage> Createonlinemeetingandthenpostmessageaboutit(string userid,string teamid,string channelid)
        {

            var onlineMeeting = new OnlineMeeting
            {
                StartDateTime = DateTimeOffset.Parse("2019-07-12T21:30:34.2444915+00:00"),
                EndDateTime = DateTimeOffset.Parse("2019-07-12T22:00:34.2464912+00:00"),
                Subject = "User Token Meeting"
            };

            var result=  await GraphServiceClient.Users[userid].OnlineMeetings
                .Request()
                .AddAsync(onlineMeeting);

           var result2= await   CreateChannelMessage("you have new assignment", channelid, teamid,userid);
            return result2;
           
        }
        public async Task AddAppsToTeamOrChannel(string teamid,string channelid)
        {
            var teamsAppInstallation = new TeamsAppInstallation
            {
                AdditionalData = new Dictionary<string, object>()
                {
                     {"teamsApp@odata.bind", "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/ef56c0de-36fc-4ef8-b417-3d82ba9d073c"}
                }
            };

            await GraphServiceClient.Teams[teamid].InstalledApps
                .Request()
                .AddAsync(teamsAppInstallation);

            //await GraphServiceClient.Teams[teamid].InstalledApps.Request().AddAsync(teamsAppInstallation);

        }
        public async Task<ICalendarEventsCollectionPage> GetUserEvents(string Id)
        {
            var events = await GraphServiceClient.Users[Id].Calendar.Events.Request().GetAsync();

            return events;
        }
        public async Task DeleteUserEvent(string userid,string EventId)
        {

            await GraphServiceClient.Users[userid].Events[EventId]
                .Request()
                .DeleteAsync();
        }
        public async Task<Event> UpdateUserEvent(string userid, string EventId, string subject)
        {


            var @event = new Event
            {
                
              
                Subject =subject,

                End = new DateTimeTimeZone
                {
                    DateTime = "2021-11-25T14:00:00",
                    TimeZone = "Pacific Standard Time"
                },

            };

           return  await GraphServiceClient.Users[userid].Events[EventId]
                .Request()
                .UpdateAsync(@event);

        }

        public async Task CreateNotification(string Id ,string Text, string teamid, string content, string messageid,string weburl)

        {
            var topic = new TeamworkActivityTopic
            {
                // 1-weburl :link chat   2 - value :klam betkb kda fl activity w feha content w title 3 - text title
                //  Source = TeamworkActivityTopicSource.EntityUrl,
                //  WebUrl = "https://teams.microsoft.com/l/message/19%3A-AbY00On0Krqn8w2FzJp6KzWibuKH_ZqGuMjS2lPivQ1%40thread.tacv2/1637501548750?groupId=e568b2b8-afd8-4599-be63-7b86e9011a66&tenantId=696e8229-c45d-433c-bee6-b6dc3377f46a&createdTime=1637501548750&parentMessageId=1637501548750",
                //Value = $"belal link ,,post link",
                Source = TeamworkActivityTopicSource.Text,
                 WebUrl =weburl,
                Value= "New Task Created"
                // Value = $"https://graph.microsoft.com/v1.0/teams/{teamid}"

                //    Source=TeamworkActivityTopicSource.EntityUrl,
                // hna weburl da optional 3lshn entityurl w l7d delw2y mrfsh lazmto eh ana momken create notificaton llmessage
                //3n tre2 lvalue w t7t tdelo lchat id 

                //  WebUrl= "https://teams.microsoft.com/l/message/19:OtcmX7flxPNssqarV590HyZ4Qvpd_FuYr6VT0mnqhRU1@thread.tacv2/1636893344822?tenantId=696e8229-c45d-433c-bee6-b6dc3377f46a&groupId=90ce8d24-bc61-4400-8039-6384f04c4eb3&parentMessageId=1636893344822&teamName=belal&channelName=General&createdTime=1636893344822",
                //   Value = $"https://graph.microsoft.com/v1.0/chats/{chatid}/messages/{messageid}",




            };

            var activityType = "reservationUpdated";

            var previewText = new ItemBody
            {
                Content = content
            };

            var recipient = new AadUserNotificationRecipient
            {

                UserId = Id
            };

            var templateParameters = new List<Microsoft.Graph.KeyValuePair>()
{
    new Microsoft.Graph.KeyValuePair
    {
        Name = "notification",
        Value = Text
    }
};

            //await GraphServiceClient.Chats[chatid]
            //    .SendActivityNotification(topic, activityType, null, previewText, templateParameters,recipient)
            //    .Request()
            //    .PostAsync();
            await GraphServiceClient.Teams[teamid]
               .SendActivityNotification(topic, activityType, null, previewText, templateParameters, recipient)
               .Request()
               .PostAsync();

            //   await GraphServiceClient.Users["7d620b50-d52d-4bb0-8c82-fcfc95ac8a12"].Teamwork
            //.SendActivityNotification(topic, activityType, null, previewText, templateParameters)
            //.Request()
            //.PostAsync();

        }


        public async Task CreateNotification2(string Id, string Text, string teamid, string content, string chatid, string messageid)
        {
            var topic = new TeamworkActivityTopic
            {
        

                Source = TeamworkActivityTopicSource.EntityUrl,
                Value = $"https://graph.microsoft.com/v1.0/chats/{chatid}/messages/{messageid}",



            };

            var activityType = "reservationUpdated";

            var previewText = new ItemBody
            {
                Content = content
            };

            var recipient = new AadUserNotificationRecipient
            {

                UserId = Id
            };



            var templateParameters = new List<Microsoft.Graph.KeyValuePair>()
{
    new Microsoft.Graph.KeyValuePair
    {
        Name = "notification",
        Value = Text
    }
};

            await GraphServiceClient.Chats[chatid]
                .SendActivityNotification(topic, activityType, null, previewText, templateParameters, recipient)
                .Request()
                .PostAsync();


        }
        public async Task<IEducationRootClassesCollectionPage> GetEducationalClasses(string externalclassid)
        {
            return await GraphServiceClient.Education.Classes
           .Request().GetAsync();
            //var filterString = $"startswith(displayName, '{externalclassid}')";
            //return await GraphServiceClient.Education.Classes
            //.Request().Filter(filterString)
            // .GetAsync();
        }
        public async Task<ITeamChannelsCollectionPage> GetDefaultChannel(string TeamId)
        {
            var filterString = $"startswith(DisplayName, 'General')";

            return await GraphServiceClient.Teams[TeamId].Channels
             .Request().Filter(filterString)
             .GetAsync();
        }
        public async Task<EducationClass> CreateEducationalClass()
        {

            var educationClass = new EducationClass
            {
                DisplayName = "MathClass",
                MailNickname = "MathClassNickName",
                Description = "any description for math class",
                CreatedBy = new IdentitySet
                {
                },
                ClassCode = "01014178103",
                ExternalName = "MathExrernalName",
                ExternalId = "7600",
                ExternalSource = EducationExternalSource.Sis,
                ExternalSourceDetail = "MathExternalSource",
                Grade = "VeryGood",
                Term = new EducationTerm
                {
                    DisplayName="Belal"
                }
            };

           return  await GraphServiceClient.Education.Classes
                .Request()
                .AddAsync(educationClass);
        }



    }
}
