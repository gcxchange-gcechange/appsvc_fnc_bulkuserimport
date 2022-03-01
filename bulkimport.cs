using System;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using System.Net.Mail;
using System.Net;
using System.Text;

namespace appsvc_fnc_dev_bulkuserimport
{
    public static class CreateUser
    {
        public static class Globals
        {
            //Global class so other class can access variables
            static IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .AddEnvironmentVariables()
            .Build();
            public static readonly string welcomeGroup = config["welcomeGroup"];
            public static readonly string GCX_Assigned = config["gcxAssigned"];
            public static readonly string UserSender = config["DoNotReplyEmail"];
            public static readonly string smtp_port = config["smtp_port"];

            private static readonly string smtp_link = config["smtp_link"];
            private static readonly string smtp_username = config["smtp_username"];
            private static readonly string smtp_password = config["smtp_password"];

            public static string GetSMTP_link()
            {
                return smtp_link;
            }
            public static string GetSMTP_username()
            {
                return smtp_username;
            }
            public static string GetSMTP_password()
            {
                return smtp_password;
            }
        }

        [FunctionName("CreateUser")]
        public static async Task RunAsync([QueueTrigger("bulkimportuserlist")] BulkInfo bulk, ILogger log)
        {

            log.LogInformation("C# HTTP trigger function processed a request.");

            string listID = bulk.listID;
            string siteID = bulk.siteID;
            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);
            log.LogInformation($"Get list id {listID}");

            var result = await getListItems(graphAPIAuth, listID, siteID, log);

            if (result == false)
            {
                throw new SystemException("Error");
            }
        }

        public static async Task<bool> getListItems(GraphServiceClient graphServiceClient, string listID, string siteID, ILogger log)
        {
            bool result = false;
            IListItemsCollectionPage list = new ListItemsCollectionPage();
            List<UsersList> userList = new List<UsersList>();
            try
            {
                var queryOptions = new List<QueryOption>()
                    {
                    //field_1 = FirstName
                    //field_2 = LastName
                    //field_3 = EmailCloud
                    //field_4 = WorkEmail
                    //field_5 = Status
                        new QueryOption("expand", "fields(select=field_1,field_2,field_3,field_4,field_5)")
                    };
                list = await graphServiceClient.Sites[siteID].Lists[listID].Items
                                        .Request(queryOptions)
                                        .GetAsync();

                var listItems = new List<ListItem>();
                listItems.AddRange(list.CurrentPage);
                while (list.NextPageRequest != null)
                {
                    list = await list.NextPageRequest.GetAsync();
                    listItems.AddRange(list.CurrentPage);
                }
                log.LogInformation(@"{count}", listItems.Count);

                foreach (var item in listItems)
                {
                    log.LogInformation(item.Fields.AdditionalData["field_2"].ToString());
                    log.LogInformation(item.Id);

                    //If status is not pending this mean it already run, don't run it again
                    if (item.Fields.AdditionalData["field_5"].ToString() == "Pending" 
                        && item.Fields.AdditionalData["field_1"].ToString() != "" 
                        && item.Fields.AdditionalData["field_2"].ToString() != "" 
                        && item.Fields.AdditionalData["field_3"].ToString() != "" 
                        && item.Fields.AdditionalData["field_4"].ToString() != "")
                    {
                        log.LogInformation(item.Fields.AdditionalData["field_2"].ToString());
                        userList.Add(new UsersList()
                        {
                            Id = item.Id,
                            FirstName = item.Fields.AdditionalData["field_1"].ToString(),
                            LastName = item.Fields.AdditionalData["field_2"].ToString(),
                            EmailCloud = item.Fields.AdditionalData["field_3"].ToString(),
                            WorkEmail = item.Fields.AdditionalData["field_4"].ToString(),
                        });

                        await UserCreation(graphServiceClient, userList, listID, siteID, log);
                        userList.Clear();
                    }
                }
                result = true;
            }
            catch (ServiceException ex)
            {
                log.LogInformation($"Error getting list : {ex.Message}");
                result = false;
            };
            return result;
        }

        public static async Task<bool> UserCreation(GraphServiceClient graphServiceClient, List<UsersList> usersList, string listID, string siteID, ILogger log)
        {
            string status = "Progress";
            string errorMessage = "";
            bool result = false;

            foreach (var item in usersList)
            {
                //update list with InProgress
                log.LogInformation("update list with progress");
                await updateList(graphServiceClient, listID, siteID, item.Id, status, errorMessage, log);

                var LastName = item.LastName;
                var FirstName = item.FirstName;
                var UserEmail = item.WorkEmail;
                bool isUserExist = true;

                // check if user exist
                try
                {
                    var userExist = await graphServiceClient.Users.Request().Filter($"mail eq '{item.EmailCloud}'").GetAsync();

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
                            InvitedUserEmailAddress = item.EmailCloud,
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
                log.LogInformation($"Error welcomegroup id : {welcomeGroup}");
                log.LogInformation($"Error assign group id : {GCX_Assigned}");

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
            string EmailSender = Globals.UserSender;
            int smtp_port = Int16.Parse(Globals.smtp_port);
            string smtp_link = Globals.GetSMTP_link();
            string smtp_username = Globals.GetSMTP_username();
            string smtp_password = Globals.GetSMTP_password();

            var Body = @$"
                        (La version française suit)<br><br>

                        Hi {FirstName} {LastName},<br><br>

                        We’re happy to announce that you now have access to gcxchange – the Government of Canada’s new digital workspace and modern intranet.<br><br>


                        Currently, there are two ways to use gcxchange: <br><br>

                        <ol><li><strong>Read articles, create and join GC-wide communities through your personalized homepage. Don’t forget to bookmark it: <a href='https://gcxgce.sharepoint.com/'>gcxgce.sharepoint.com/</a></strong></li>

                        <li><strong>Chat, call, and co-author with members of your communities using your Microsoft Teams and seamlessly toggle between gcxchange and your departmental environment. <a href='https://teams.microsoft.com/_?tenantId=f6a7234d-bc9b-4520-ad5f-70669c2c7a9c#/conversations/General?threadId=19:OXWdygF2pylAN26lrbZNN-GGzf8W9YEpe2EBawXtM0s1@thread.tacv2&ctx=channel'>Click here to find out how!</a></strong></li></ol>

                        We want to hear from you! Please take a few minutes to respond to our <a href=' https://questionnaire.simplesurvey.com/f/l/gcxchange-gcechange?idlang=EN'>survey</a> about the registration process.<br><br>

                        If you run into any issues along the way, please reach out to the support team at: <a href='mailto:support-soutien@gcx-gce.gc.ca'>support-soutien@gcx-gce.gc.ca</a><br><br>
                        
                        ---------------------------------------------------------------------------------<br><br>

                        (The English version precedes)<br><br>

                        Bonjour {FirstName} {LastName},<br><br>

                        Nous sommes heureux de vous annoncer que vous avez maintenant accès à gcéchange – le nouvel espace de travail numérique et intranet moderne du gouvernement du Canada.<br><br>


                        À l’heure actuelle, il y a deux façons d’utiliser gcéchange : <br><br>

                        <ol><li><strong>Lisez des articles, créez des communautés pangouvernementales et joignez-vous à celles-ci au moyen de votre page d’accueil personnalisée. N’oubliez pas d’ajouter cet espace dans vos favoris : <a href='https://gcxgce.sharepoint.com/'>gcxgce.sharepoint.com/</a></strong></li>

                        <li><strong>Clavardez et corédigez des documents avec des membres de vos communautés ou appelez ces membres au moyen de Microsoft Teams et passez facilement de gcéchange à votre environnement ministériel. <a href='https://teams.microsoft.com/_?tenantId=f6a7234d-bc9b-4520-ad5f-70669c2c7a9c#/conversations/General?threadId=19:OXWdygF2pylAN26lrbZNN-GGzf8W9YEpe2EBawXtM0s1@thread.tacv2&ctx=channel'>Cliquez ici pour savoir comment faire.</a></strong></li></ol>

                        Nous souhaitons connaître votre opinion! Veuillez prendre quelques minutes pour répondre à notre <a href='https://questionnaire.simplesurvey.com/f/l/gcxchange-gcechange?idlang=FR'>sondage</a> sur le processus d’inscription.<br><br>


                        Si vous éprouvez des problèmes en cours de route, veuillez communiquer avec l’équipe de soutien à l’adresse suivante : <a href='mailto:support-soutien@gcx-gce.gc.ca'>support-soutien@gcx-gce.gc.ca</a>";

            MailMessage mail = new MailMessage();

            mail.From = new MailAddress(EmailSender);
            mail.To.Add(UserEmail);
            mail.Subject = "You're in! | Vous s'y êtes";
            mail.Body = Body;
            mail.IsBodyHtml = true;

            SmtpClient SmtpServer = new SmtpClient(smtp_link);
            SmtpServer.Port = smtp_port;
            SmtpServer.Credentials = new System.Net.NetworkCredential(smtp_username, smtp_password);
            SmtpServer.EnableSsl = true;

            log.LogInformation($"UserEmail : {UserEmail}");

            try
            {
                SmtpServer.Send(mail);
                log.LogInformation("mail Send");
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
            //field_5 = status
            //field_6 = ErrorMessage
            var fieldValueSet = new FieldValueSet
            {                    
                AdditionalData = new Dictionary<string, object>()
                {
                    {"field_5", status},
                    {"field_6", errorMessage}
                }
            };

            try
            {
                await graphServiceClient.Sites[siteID].Lists[listID].Items[itemID].Fields
                    .Request()
                    .UpdateAsync(fieldValueSet);
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error list ids : {listID}");
                log.LogInformation($"Error site id : {siteID}");
                log.LogInformation($"Error item id : {itemID}");

                log.LogInformation($"Error updating the list: {ex}");
            }
            return true;
        }
    }
}
