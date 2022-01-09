using Azure.Identity;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;

namespace MS.Teams.POC.WebApp.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public async Task<IActionResult> ListAllApplications()
        {
            try
            {
              
                var ProtectedApiCallHelper = await RunAsync(true);

                var res = (await ProtectedApiCallHelper.GetApplicationAsync());
                return Ok(res);
            }
            catch (Exception ex)
            {
                var x = ex.Message;
                throw;
            }
        }
        [HttpGet]
        public async Task<IActionResult> ListAllChats(string SendTo)
        {
            try
            {
                //delegate permissions required 
                //Chat.ReadBasic, Chat.Read, Chat.ReadWrite
                var ProtectedApiCallHelper = await RunAsync(false);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

                var res = (await ProtectedApiCallHelper.ListAllChats(user.Id));
                return Ok(res);
            }
            catch (Exception ex)
            {
                var x = ex.Message;
                throw;
            }
        }
        [HttpPost]
        public async Task<IActionResult> CreateonetooneChat(string SendTo)
        {
            try
            {
                //delegate permission required
                //Chat.Create, Chat.ReadWrite
                var ProtectedApiCallHelper = await RunAsync(false);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

                var res = (await ProtectedApiCallHelper.CreateOneToOneChat(user.Id));
                return Ok(res);
            }
            catch (Exception ex)
            {
                var x = ex.Message;
                throw;
            }
        }
        [HttpPost]
        public async Task<IActionResult> CreatechannelMessage(string content, string channelid,string teamid,string SendTo)
        {
            try
            {
                //delegate permission required
                //Chat.Create, Chat.ReadWrite
                var ProtectedApiCallHelper = await RunAsync(false);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

                await ProtectedApiCallHelper.CreateChannelMessage(content, channelid, teamid, user.Id);
                ViewBag.Data = "Done";
            }
            catch (Exception ex)
            {
                var x = ex.Message;
                throw;
            }
            return View("Create");
        }
        [HttpPost]
        public async Task<IActionResult> CreateNotification(string Title, string SendTo,string teamid,string content,string chatid,string messageid,string weburl)
        {
            ViewBag.Title = Title;

            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);
               

                if (user != null)
                {
                    await ProtectedApiCallHelper.CreateNotification(user.Id, Title,teamid,content,messageid, weburl);
                    ViewBag.Data = "Done";
                }
                else
                {
                    ViewBag.Data = "Failed";
                }
            }
            catch (Exception ex)
            {
                ViewBag.Data = "Failed";
                throw;
            }
            return View("Create");
        }
        [HttpPost]
        public async Task<IActionResult> CreateNotification2(string Title, string SendTo, string teamid, string content, string chatid, string messageid)
        {
            ViewBag.Title = Title;

            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);


                if (user != null)
                {
                    await ProtectedApiCallHelper.CreateNotification2(user.Id, Title, teamid, content, chatid, messageid);
                    ViewBag.Data = "Done";
                }
                else
                {
                    ViewBag.Data = "Failed";
                }
            }
            catch (Exception ex)
            {
                ViewBag.Data = "Failed";
                throw;
            }
            return View("Create");
        }
        [HttpGet]
        public async Task<IActionResult> ListUserCalendars(string SendTo)
        {
            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

                if (user != null)
                {
                    return Ok(await ProtectedApiCallHelper.ListUserCalendrs(SendTo));
                }
                return NotFound();
            }
            catch (Exception ex)
            {
                var x = ex.Message;

                throw;
            }

        }
        [HttpPost]
        public async Task<IActionResult> DeleteCalendar(string calendarid,string SendTo)
        {
            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

                if (user != null)
                {
                   await ProtectedApiCallHelper.DeleteCalendar(calendarid, user.Id);
                    return Ok();
                }
                return NotFound();
            }
            catch (Exception ex)
            {
                var x = ex.Message;

                throw;
            }
        }

            
        [HttpGet]
        public async Task<IActionResult> GetAllchannels(string teamid,string SendTo)
        {
            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

                if (user != null)
                {
                    return Ok(await ProtectedApiCallHelper.GetChannels(teamid));
                }
                return NotFound();
            }
            catch(Exception ex)
            {
                var x = ex.Message;

                throw;
            }

        }
        [HttpGet]
         public async Task<IActionResult> GetAllteams(string SendTo)
        {
            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

                if (user != null)
                {
                    return Ok(await ProtectedApiCallHelper.GetTeams(user.Id));
                }
                return NotFound();
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        [HttpPost]
        public async Task<IActionResult> CreateCalendar(string Title, string SendTo)
        {
         
            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);
                if (user != null)
                {
                    ViewBag.Data = await ProtectedApiCallHelper.CreateUserCalendar(user.Id,Title);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return View("Create");
        }
        [HttpPost]
        public async Task<IActionResult> DeleteUserEvent(string SendTo,string EventId)
        {
            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);
                if (user != null)
                {
                    await ProtectedApiCallHelper.DeleteUserEvent(user.Id, EventId);
                    return Ok();
                }
                return NotFound();
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        [HttpPost]
        public async Task<IActionResult> UpdateUserEvent(string SendTo, string EventId,string subject)
        {
            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);
                if (user != null)
                {
                 return   Ok(await ProtectedApiCallHelper.UpdateUserEvent(user.Id, EventId, subject));
                  
                }
                return NotFound();
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        [HttpPost]
        public async Task<IActionResult> AddmemberToTeam(string teamid ,string memberid,string azureuserid)

        {
            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(azureuserid);

             await ProtectedApiCallHelper.AddMemberToTeam(memberid,teamid);
                return Ok();
                
            }
            catch (Exception ex)
            {
                var x = ex.Message;

                throw;
            }
        }
         [HttpGet]
         public async Task<IActionResult> GetAllmessagesinchannel(string teamid,string channelid,string SendTo)
        {
            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

                var res = (await ProtectedApiCallHelper.GetAllMessagesAtChannel(teamid, channelid));
                return Ok(res);
            }
            catch (Exception ex)
            {
                var x = ex.Message;

                throw;
            }
        }
        
        [HttpPost]
        public async Task<IActionResult> CreateTeam(string DisplayName, string Description,string SendTo)
        {
            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

                var res = (await ProtectedApiCallHelper.CreateTeam(DisplayName, Description));
                return Ok(res);
            }
            catch (Exception ex)
            {
                var x = ex.Message;

                throw;
            }
        }

        [HttpPost]
        public async Task<IActionResult> CreateChannel(string DisplayName,string Description,string id,string SendTo)
        {
            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

              var res= (await ProtectedApiCallHelper.CreateChannel(DisplayName, Description, id, user.Id));
                return Ok(res);
            }
            catch (Exception ex)
            {
                var x = ex.Message;

                throw;
            }
        }
      
        [HttpPost]
        public async Task<IActionResult> CreateEvent(string Title, string Description, string SendTo,string calendarname)
        {
            ViewBag.Title = Title;
            ViewBag.Description = Description;

            try
            {
                //delegate permission required

                // Group.ReadWrite.All
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

                if (user != null)
                {
                    ViewBag.Data = await ProtectedApiCallHelper.CreateUserEvent(user.Id, calendarname, GetEvent(Title, Description, SendTo));
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return View("Create");
        }
        [HttpPost]
        public async Task<IActionResult> Createonlinemeetingandthenpostmessageaboutit(string SendTo,string teamid,string channelid)
        {
            try
            {

                var ProtectedApiCallHelper = await RunAsync(false);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);
                if (user != null)
                {
                    await ProtectedApiCallHelper.Createonlinemeetingandthenpostmessageaboutit(user.Id,teamid,channelid);
                    return Ok();
                }
                return NotFound();
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        [HttpPost]
        public async Task<IActionResult> AddAppsToTeamOrChannel(string teamid,string channelid,string SendTo)
        {
            try
            {
              
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);
                if (user != null)
                {
                     await ProtectedApiCallHelper.AddAppsToTeamOrChannel(teamid,channelid);
                    return Ok();
                }
                return NotFound();     
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private static Event GetEvent(string Title, string Description, string SendTo)
        {
            return new Event
            {
                Subject = Title,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = Description
                },
                Start = new DateTimeTimeZone
                {
                    DateTime = DateTime.UtcNow.AddDays(2).ToString(),
                    TimeZone = "UTC"
                },
                End = new DateTimeTimeZone
                {
                    DateTime = DateTime.UtcNow.AddDays(5).ToString(),
                    TimeZone = "UTC"
                },
                Location = new Location
                {
                    DisplayName = Title
                },
                Attendees = new List<Attendee>()
                {
                    new Attendee
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = SendTo,
                           Name = SendTo
                        },
                        Type = AttendeeType.Required
                    }
                },
               
                TransactionId = Guid.NewGuid().ToString()
            };
        }
    
        private async Task<ProtectedApiCallHelper> RunAsync(bool IsApplicationPermission)
        {
            AuthenticationConfig config = AuthenticationConfig.ReadFromJsonFile("appsettings.json");

            // You can run this sample using ClientSecret or Certificate. The code will differ only when instantiating the IConfidentialClientApplication
            bool isUsingClientSecret = AppUsesClientSecret(config);
            AuthenticationResult result = null;

            if (IsApplicationPermission)
            {
                // Even if this is a console application here, a daemon application is a confidential client application
                IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(config.ClientId)
                        .WithClientSecret(config.ClientSecret)
                        .WithAuthority(new Uri(config.Authority))

                        .Build();



                // With client credentials flows the scopes is ALWAYS of the shape "resource/.default", as the 
                // application permissions need to be set statically (in the portal or by PowerShell), and then granted by
                // a tenant administrator. 
                string[] scopes = new string[] { $"{config.ApiUrl}.default" };

                try
                {
                    result = await app.AcquireTokenForClient(scopes)
                        .ExecuteAsync();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Token acquired");
                    Console.ResetColor();
                }

                catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS70011"))
                {
                    // Invalid scope. The scope has to be of the form "https://resourceurl/.default"
                    // Mitigation: change the scope to be as expected
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Scope provided is not supported");
                    Console.ResetColor();
                }

            }
            else
            {
                // begin trying for authenticate user


                string[] scopes = new string[] { "user.read" };
                IPublicClientApplication app;
                app = PublicClientApplicationBuilder.Create(config.ClientId)
                      .WithAuthority(new Uri(config.Authority))
                      .Build();
                var accounts = await app.GetAccountsAsync();
                if (accounts.Any())
                {
                    result = await app.AcquireTokenSilent(scopes, accounts.Where(x => x.Username == "admin.teams@edu-worx.net").FirstOrDefault())
                                      .ExecuteAsync();
                }
                else
                {
                    try
                    {
                        var securePassword = new SecureString();
                        foreach (char c in config.Password)        // you should fetch the password
                            securePassword.AppendChar(c);  // keystroke by keystroke

                        result = await app.AcquireTokenByUsernamePassword(scopes,
                                                                         config.username,
                                                                          securePassword)
                                           .ExecuteAsync();
                    }
                    catch (MsalException)
                    {
                        // See details below
                    }
                }
                // end trying of authenticate user

            }


            if (result != null)
            {
                var authProvider = new DelegateAuthenticationProvider(async (request) =>
                {
                    request.Headers.Authorization =
                        new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer",result.AccessToken);
                });
                var graphClient = new GraphServiceClient(authProvider);
                var graphCaller = new ProtectedApiCallHelper(graphClient);
                return graphCaller;
            }
            return null;
        }

        /// <summary>
        /// Checks if the sample is configured for using ClientSecret or Certificate. This method is just for the sake of this sample.
        /// You won't need this verification in your production application since you will be authenticating in AAD using one mechanism only.
        /// </summary>
        /// <param name="config">Configuration from appsettings.json</param>
        /// <returns></returns>
        private bool AppUsesClientSecret(AuthenticationConfig config)
        {
            string clientSecretPlaceholderValue = "[Enter here a client secret for your application]";
            string certificatePlaceholderValue = "[Or instead of client secret: Enter here the name of a certificate (from the user cert store) as registered with your application]";

            if (!String.IsNullOrWhiteSpace(config.ClientSecret) && config.ClientSecret != clientSecretPlaceholderValue)
            {
                return true;
            }

            else if (!String.IsNullOrWhiteSpace(config.CertificateName) && config.CertificateName != certificatePlaceholderValue)
            {
                return false;
            }

            else
                throw new Exception("You must choose between using client secret or certificate. Please update appsettings.json file.");
        }

        [HttpPost]
        public async Task<IActionResult> GetUserEvents(string userEmail)
        {
            
            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(userEmail);

                if (user != null)
                {
                    ViewBag.Data = await ProtectedApiCallHelper.GetUserEvents(user.Id);

                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return View("Create");
        }
        [HttpGet]
        public async Task<IActionResult> GetEducationalClasses(string SendTo)
        {

            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

                if (user != null)
                {
                    ViewBag.Data = await ProtectedApiCallHelper.GetEducationalClasses("Grade 4");

                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return View("Create");
        }
        [HttpGet]
        public async Task<IActionResult> GetDefaultChannel(string SendTo,string TeamId)
        {

            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

                if (user != null)
                {
                    ViewBag.Data = await ProtectedApiCallHelper.GetDefaultChannel(TeamId);

                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return View("Create");
        }
        [HttpPost]
        public async Task<IActionResult> CreateEducationalClass(string SendTo)
        {

            try
            {
                var ProtectedApiCallHelper = await RunAsync(true);
                var user = await ProtectedApiCallHelper.GetUser(SendTo);

                if (user != null)
                {
                    ViewBag.Data = await ProtectedApiCallHelper.CreateEducationalClass();

                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return View("Create");
        }

    }
}
