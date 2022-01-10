using System;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;

namespace appsvc_fnc_dev_bulkuserimport
{
    public static class CreateUser
    {
        public static class Globals
        {
            static IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .AddEnvironmentVariables()
            .Build();
            public static readonly string welcomeGroup = config["welcomeGroup"];
            public static readonly string GCX_Assigned = config["gcxAssigned"];
            public static readonly string UserSender = config["UserSender"];

        }
        [FunctionName("CreateUser")]
        public static async Task RunAsync(
            [QueueTrigger("bulkimportuserlist")] BulkInfo bulk,
            ILogger log)
        {

            log.LogInformation("C# HTTP trigger function processed a request.");

            string listID = bulk.listID;
            string siteID = bulk.siteID;
            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);
            log.LogInformation($"Get list id {listID}");

            var result = await getListItems(graphAPIAuth, listID, siteID, log);

            // var createUser = await UserCreation(graphAPIAuth, EmailCloud, FirstName, LastName, redirectLink, log);

            if (result == "false")
            {
                throw new SystemException("Error");
            }
        }

        public static async Task<string> getListItems(GraphServiceClient graphServiceClient, string listID, string siteID, ILogger log)
        {
            //List<string> list = new List<string>();
            //  var list = new IListItemsCollectionPage();
            IListItemsCollectionPage list = new ListItemsCollectionPage();
            List<UsersList> userList = new List<UsersList>();
            try
            {
                var queryOptions = new List<QueryOption>()
                    {
                        new QueryOption("expand", "fields(select=FirstName,LastName,DepartmentEmail,WorkEmail)")
                    };
                list = await graphServiceClient.Sites[siteID].Lists[listID].Items
                        .Request(queryOptions)
                        .GetAsync();

                foreach (var item in list)
                {
                    var FirstName = item.Fields.AdditionalData["FirstName"];
                    log.LogInformation($"{FirstName}");
                    var LastName = item.Fields.AdditionalData["LastName"];
                    log.LogInformation($"{LastName}");
                    log.LogInformation("Name:" + item.Name);
                    log.LogInformation($"ID: {item.Id}");
                    userList.Add(new UsersList()
                    {
                        Id = item.Id,
                        FirstName = item.Fields.AdditionalData["FirstName"].ToString(),
                        LastName = item.Fields.AdditionalData["LastName"].ToString(),
                        DepartmentEmail = item.Fields.AdditionalData["DepartmentEmail"].ToString(),
                        WorkEmail = item.Fields.AdditionalData["WorkEmail"].ToString()
                    });

                    await UserCreation(graphServiceClient, userList, listID, siteID, log);

                    userList.Clear();
                }


                //return list;
            }
            catch (ServiceException ex)
            {
                log.LogInformation($"Error getting list : {ex.Message}");
                //return "something";
                //InviteInfo.Add("Invitation error");
            };

            //return await Task.FromResult("true");
            return "yes";
        }

        public static async Task<bool> UserCreation(GraphServiceClient graphServiceClient, List<UsersList> usersList, string listID, string siteID, ILogger log)
        {
            string status = "Progress";
            string errorMessage = "";
            bool result = false;
            log.LogInformation("in usercreation");
            //List<string> InviteInfo = new List<string>();
            //IListItemsCollectionPage list = usersList;
            foreach (var item in usersList)
            {
                //update list with InProgress
                log.LogInformation("update list with progress");
                await updateList(graphServiceClient, listID, siteID, item.Id, status, errorMessage, log);

                var LastName = item.LastName;
                var FirstName = item.FirstName;
                var UserEmail = item.WorkEmail;
                log.LogInformation($"{LastName}");
                log.LogInformation($"{item.Id}");
                bool isUserExist = true;
                // check if user exist
                try
                {
                    var userExist = await graphServiceClient.Users.Request().Filter($"mail eq '{item.WorkEmail}'").GetAsync();
                    if (userExist.Count > 0)
                    {
                        isUserExist = true;
                        log.LogInformation($"User exist");
                        status = "Error";
                        errorMessage = $"User already exist";
                        await updateList(graphServiceClient, listID, siteID, item.Id, status, errorMessage, log);
                        isUserExist = true;
                    }
                    else
                    {
                        isUserExist = false;
                    }
                }
                catch (Exception ex)
                {
                    log.LogInformation($"error checking if user exist : {ex.Message}");
                    status = "Error";
                    errorMessage = $"User already exist {ex.Message}";
                    await updateList(graphServiceClient, listID, siteID, item.Id, status, errorMessage, log);
                    isUserExist = true;
                    result = false;
                }


                if (isUserExist == false)
                {
                    var userinviteID = "";
                    //Create Invitation
                    try
                    {
                        var invitation = new Invitation
                        {
                            SendInvitationMessage = false,
                            InvitedUserEmailAddress = item.WorkEmail,
                            InvitedUserType = "member",
                            InviteRedirectUrl = "https://gcxgce.sharepoint.com/",
                            InvitedUserDisplayName = $"{item.FirstName} {item.LastName}",
                        };

                        var userinvite = await graphServiceClient.Invitations.Request().AddAsync(invitation);
                        userinviteID = userinvite.InvitedUser.Id;
                        log.LogInformation($"user invite successfully - {userinvite.InvitedUser.Id}");
                    }
                    catch (Exception ex)
                    {
                        log.LogInformation($"error creating user invite : {ex.Message}");
                        status = "Error";
                        errorMessage = ex.Message;
                        await updateList(graphServiceClient, listID, siteID , item.Id, status, errorMessage, log);
                        result = false;
                    };

                    if (userinviteID != "")
                    {
                        bool updateuser = false;
                        try
                        {
                            updateuser = await updateUser(graphServiceClient, userinviteID, log);
                        }
                        catch (Exception ex)
                        {
                            updateuser = false;
                            log.LogInformation($"Error Updating User : {ex.Message}");
                            status = "Error";
                            errorMessage = ex.Message;
                            await updateList(graphServiceClient, listID, siteID, item.Id, status, errorMessage, log);
                            result = false;
                        }

                        if (updateuser)
                        {
                            //add user to all group
                            bool addGroup = false;
                            try
                            {
                                addGroup = await addUserToGroups(graphServiceClient, userinviteID, siteID, listID, item.Id, log);
                            }
                            catch (Exception ex)
                            {
                                addGroup = false;
                                log.LogInformation($"Error adding user to groups : {ex.Message}");
                                status = "Error";
                                errorMessage = ex.Message;
                                await updateList(graphServiceClient, listID, siteID, item.Id, status, errorMessage, log);
                                result = false;
                            }

                            if (addGroup)
                            {
                                //add user to all group
                                bool sendEmail = false;
                                try
                                {
                                    sendEmail = await sendUserEmail(graphServiceClient, FirstName, LastName, UserEmail, listID, siteID, item.Id, log);
                                    result = true;
                                }
                                catch (Exception ex)
                                {
                                    sendEmail = false;
                                    log.LogInformation($"Error sending email to user : {ex.Message}");
                                    status = "Error";
                                    errorMessage = ex.Message;
                                    await updateList(graphServiceClient, listID, siteID, item.Id, status, errorMessage, log);
                                    result = false;
                                }

                                if (sendEmail)
                                {
                                    //Update SP list
                                    log.LogInformation($"Update sp list with completed");
                                    status = "completed";
                                    errorMessage = "";
                                    await updateList(graphServiceClient, listID, siteID, item.Id, status, errorMessage, log);
                                    result = true;

                                }
                            }
                        }
                    }
                }
            }
            return result;
        }

        public static async Task<bool> updateUser(GraphServiceClient graphServiceClient, string userID, ILogger log)
        {
            bool result = false;
            try
            {
                var guestUser = new User
                {
                    UserType = "Member"
                };

                await graphServiceClient.Users[userID].Request().UpdateAsync(guestUser);
                log.LogInformation("User update successfully");

                result = true;
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error Updating User : {ex.Message}");
                result = false;
            }
            return result;
        }

        public static async Task<bool> addUserToGroups(GraphServiceClient graphServiceClient, string userID, string listID, string siteID, string itemID, ILogger log)
        {
            bool result = false;
            string welcomeGroup = Globals.welcomeGroup;
            string GCX_Assigned = Globals.GCX_Assigned;
            try
            {
                var directoryObject = new DirectoryObject
                {
                    Id = userID
                };

                await graphServiceClient.Groups[welcomeGroup].Members.References
                    .Request()
                    .AddAsync(directoryObject);
                log.LogInformation("User add to welcome group successfully");
                await graphServiceClient.Groups[GCX_Assigned].Members.References
                    .Request()
                    .AddAsync(directoryObject);
                log.LogInformation("User add to GCX_Assigned group successfully");

                result = true;
            }
            catch (Exception ex)
            {
                //var listID = CreateUser.listID
                log.LogInformation($"Error adding User groups : {ex.Message}");
                string status = "Error";
                string errorMessage = ex.Message;
                await updateList(graphServiceClient, listID, siteID, itemID, status, errorMessage, log);
                result = false;
            }
            return result;
        }

        public static async Task<bool> sendUserEmail(GraphServiceClient graphServiceClient, string FirstName, string LastName, string UserEmail, string listID, string siteID, string itemID, ILogger log)
        {
            bool result = false;
            var submitMsg = new Message();
            string EmailSender = Globals.UserSender;
            submitMsg = new Message
            {
                Subject = "You're in! | Vous s'y �tes",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = @$"
                        (La version fran�aise suit)<br><br>

                        Hi {FirstName} {LastName},<br><br>

                        We�re happy to announce that you now have access to gcxchange � the Government of Canada�s new digital workspace and modern intranet.<br><br>


                        Currently, there are two ways to use gcxchange: <br><br>

                        <ol><li><strong>Read articles, create and join GC-wide communities through your personalized homepage. Don�t forget to bookmark it: <a href='https://gcxgce.sharepoint.com/'>gcxgce.sharepoint.com/</a></strong></li>

                        <li><strong>Chat, call, and co-author with members of your communities using your Microsoft Teams and seamlessly toggle between gcxchange and your departmental environment. <a href='https://teams.microsoft.com/_?tenantId=f6a7234d-bc9b-4520-ad5f-70669c2c7a9c#/conversations/General?threadId=19:OXWdygF2pylAN26lrbZNN-GGzf8W9YEpe2EBawXtM0s1@thread.tacv2&ctx=channel'>Click here to find out how!</a></strong></li></ol>

                        We want to hear from you! Please take a few minutes to respond to our <a href=' https://questionnaire.simplesurvey.com/f/l/gcxchange-gcechange?idlang=EN'>survey</a> about the registration process.<br><br>

                        If you run into any issues along the way, please reach out to the support team at: <a href='mailto:support-soutien@gcx-gce.gc.ca'>support-soutien@gcx-gce.gc.ca</a><br><br>
                        
                        ---------------------------------------------------------------------------------<br><br>

                        (The English version precedes)<br><br>

                        Bonjour {FirstName} {LastName},<br><br>

                        Nous sommes heureux de vous annoncer que vous avez maintenant acc�s � gc�change � le nouvel espace de travail num�rique et intranet moderne du gouvernement du Canada.<br><br>


                        � l�heure actuelle, il y a deux fa�ons d�utiliser gc�change : <br><br>

                        <ol><li><strong>Lisez des articles, cr�ez des communaut�s pangouvernementales et joignez-vous � celles-ci au moyen de votre page d�accueil personnalis�e. N�oubliez pas d�ajouter cet espace dans vos favoris : <a href='https://gcxgce.sharepoint.com/'>gcxgce.sharepoint.com/</a></strong></li>

                        <li><strong>Clavardez et cor�digez des documents avec des membres de vos communaut�s ou appelez ces membres au moyen de Microsoft Teams et passez facilement de gc�change � votre environnement minist�riel. <a href='https://teams.microsoft.com/_?tenantId=f6a7234d-bc9b-4520-ad5f-70669c2c7a9c#/conversations/General?threadId=19:OXWdygF2pylAN26lrbZNN-GGzf8W9YEpe2EBawXtM0s1@thread.tacv2&ctx=channel'>Cliquez ici pour savoir comment faire.</a></strong></li></ol>

                        Nous souhaitons conna�tre votre opinion! Veuillez prendre quelques minutes pour r�pondre � notre <a href='https://questionnaire.simplesurvey.com/f/l/gcxchange-gcechange?idlang=FR'>sondage</a> sur le processus d�inscription.<br><br>


                        Si vous �prouvez des probl�mes en cours de route, veuillez communiquer avec l��quipe de soutien � l�adresse suivante : <a href='mailto:support-soutien@gcx-gce.gc.ca'>support-soutien@gcx-gce.gc.ca</a>"
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                           Address = $"{UserEmail}"
                        }
                    }
                },
            };
            try
            {
                await graphServiceClient.Users[EmailSender]
                      .SendMail(submitMsg)
                      .Request()
                      .PostAsync();
                log.LogInformation($"User mail successfully");
                result = true;
            }
            
            catch (ServiceException ex)
            {
                log.LogInformation($"Error sending email: {ex.Message}");
                string status = "Error";
                string errorMessage = ex.Message;
                await updateList(graphServiceClient, listID, siteID, itemID, status, errorMessage, log);
                result = false;
            }

            return result;
        }


    public static async Task<bool> updateList(GraphServiceClient graphServiceClient, string listID, string siteID, string itemID, string status, string errorMessage, ILogger log)
        {
            var fieldValueSet = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object>()
                {
                    {"Status", status},
                    {"ErrorMessage", errorMessage}
                }
            };
            log.LogInformation($"Update item {itemID}");
            log.LogInformation($"Update list {listID}");
            log.LogInformation($"Update site {siteID}");


            try
            {
                await graphServiceClient.Sites[siteID].Lists[listID].Items[itemID].Fields
                    .Request()
                    .UpdateAsync(fieldValueSet);
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error updating the list: {ex}");
            }
            return true;
        }
    }
}