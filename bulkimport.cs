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
                        <div style='font-family: Helvetica'>
			<i>(La version française suit)</i> 
			<br><br>
			<b>Welcome to GC<b style='color: #1f9cf5'>X</b>change!</b>
			<br><br>
			Hi { FirstName } { LastName },
			<br><br>
			Your department has given you access to <a href='https://gcxgce.sharepoint.com/?gcxLangTour=en'>GCXchange</a> - the GC's new digital workspace and collaboration platform! No log-in or password is needed for GCXchange, since it uses a single sign-on from your government device.
			<br><br>
			<center><h2><a href='https://gcxgce.sharepoint.com/?gcxLangTour=en'>You can access GCXchange here</a></h2>
			<br>
			<b>Bookmark the above link to your personalized homepage, as well as to <a href='https://teams.microsoft.com/_?tenantId=f6a7234d-bc9b-4520-ad5f-70669c2c7a9c#/conversations/General?threadId=19:OXWdygF2pylAN26lrbZNN-GGzf8W9YEpe2EBawXtM0s1@thread.tacv2&ctx=channel'>GCXchange's MS Teams platform.</a></b></center>
			<br><br>
			GCXchange uses a combination of Sharepoint and MS Teams to allow users to collaborate across GC departments and agencies.
			<br><br>
			On the Sharepoint side of GCXchange you can:
			<ol>
				<li>Read <a href='https://gcxgce.sharepoint.com/sites/news'>GC-wide news and stories</a>.</li>
				<li>Join one of the many <a href='https://gcxgce.sharepoint.com/sites/Communities'>cross-departmental communities.</a></li>
				<li>Engage with thematic hubs that focus on issues relevant to the public service.</li>
				<li>Create a <a href='https://gcxgce.sharepoint.com/sites/Support/SitePages/Communities.aspx'>community</a> for interdepartmental collaboration with a dedicated page and Teams space.</li>
			</ol>
			<br>
			On the Teams side of GCXchange you can engage with the communities you have joined, as well as co-author documents and chat with colleagues in other departments and agencies. To learn how to switch between your departmental and GCXchange MS Teams accounts <a href='https://www.youtube.com/watch?v=71bULf1UqGw&list=PLWhPHFzdUwX98NKbSG8kyq5eW9waj3nNq&index=8'>watch a video tutorial</a> or <a href='https://gcxgce.sharepoint.com/sites/Support/SitePages/FAQ.aspx'>access the step-by-step guidance</a>.
			<br><br>
			If you run into a problem or have a question, contact: <a href='mailto:support-soutien@gcx-gce.gc.ca'>support-soutien@gcx-gce.gc.ca</a>
			<br><br>
			Happy collaborating!
			<br><br>
			<hr>
			<br><br>
			<b>Bienvenue à GC<b style='color: #1f9cf5'>É</b>change!</b>
			<br><br>
			Bonjour { FirstName } { LastName },
			<br><br>
			Votre ministère vous a donné accès à <a href='https://gcxgce.sharepoint.com/SitePages/fr/Home.aspx?gcxLangTour=fr'>GCÉchange</a>, la nouvelle plateforme de collaboration et de travail numérique du GC! Aucum nom d'utilisateur ni mot de passe n'est requis pour accéder à GCÉchange, puisque cette platforme est intégrée à la session unique que vous ouvrez à partir de votre appareil gouvernemental.
			<br><br>
			<center><h2><a href='https://gcxgce.sharepoint.com/SitePages/fr/Home.aspx?gcxLangTour=fr'>Vous pouvez accéder à GCÉchange ici</a></h2>
			<br>
			<b>Ajoutez le lien ci-dessus comme favori à votre page d'accueil personnalisée ainsi qu'à <a href='https://teams.microsoft.com/_?tenantId=f6a7234d-bc9b-4520-ad5f-70669c2c7a9c#/conversations/General?threadId=19:OXWdygF2pylAN26lrbZNN-GGzf8W9YEpe2EBawXtM0s1@thread.tacv2&ctx=channel'>la plateforme Microsoft Teams de gcéchange.</a></b></center>
			<br><br>
			GCÉchange utilise SharePoint et Teams pour permettre aux utilisateurs de collaborer avec l'ensemble des ministères et organismes du GC.
			<br><br>
			Du côté SharePoint de GCÉchange, vous pouvez :
			<ol>
				<li>lire <a href='https://gcxgce.sharepoint.com/sites/News/SitePages/fr/Home.aspx'></a>les nouvelles et les histoires du GC.</li>
				<li>participer à l'une des nombreuses <a href='https://gcxgce.sharepoint.com/sites/Communities/SitePages/fr/Home.aspx'>participer à l'une des nombreuses collectivités interministérielles.</a></li>
				<li>participer à des carrefours thématiques qui se concentrent sur ces enjeux pertinents pour la fonction publique.</li>
				<li>créer une <a href='https://gcxgce.sharepoint.com/sites/Support/SitePages/fr/Communities.aspx'>collectivité</a> de collaboration interministérielle qui a sa page et son espace Teams.</li>
			</ol>
			<br>
			Du côté Teams de GCÉchange, vous pouvez communiquer avec les collectivités desquelles vous êtes membre, corédiger des documents et clavarder avec des collègues d'autres ministères et organismes. Pour savoir comment passer d'un compte ministériel à un compte GCÉchange dans Teams, <a href='https://gcxgce.sharepoint.com/sites/Support/SitePages/fr/FAQ.aspx'>regardez un tutoriel vidéo ou accédez aux directives étape par étape.</a>
			<br><br>
			Si vous avez un problème ou une question, écrivez à : <a href='mailto:support-soutien@gcx-gce.gc.ca'>support-soutien@gcx-gce.gc.ca</a>.
			<br><br>
			Bonne collaboration!
			</div>";

            MailMessage mail = new MailMessage();

            mail.From = new MailAddress(EmailSender);
            mail.To.Add(UserEmail);
            mail.Subject = "Welcome to GCXchange! | Bienvenue à GCÉchange!";
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
