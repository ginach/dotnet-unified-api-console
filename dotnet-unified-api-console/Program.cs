#region

using System;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.OData.Client;
using System.Net.Http;

#endregion

namespace MicrosoftGraphSampleConsole
{
    public class Program
    {

        // Single-Threaded Apartment required for OAuth2 Authz Code flow (User Authn) to execute for this demo app
        [STAThread]
        private static void Main()
        {
            // record start DateTime of execution
            string currentDateTime = DateTime.Now.ToUniversalTime().ToString();

            ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;

            Console.WriteLine("Run operations for signed-in user, or in app-only mode.\n");
            Console.WriteLine("[a] - app-only\n[u] - as user\n[b] - both as user first, and then as app.\nPlease enter your choice:\n");

            ConsoleKeyInfo key = Console.ReadKey();
            switch (key.KeyChar)
            {
                case 'a':
                    Console.WriteLine("\nRunning app-only mode\n\n");
                    Requests.AppMode();
                    break;
                case 'b':
                    Console.WriteLine("\nRunning app-only mode, followed by user mode\n\n");
                    break;
                case 'u':
                    Console.WriteLine("\nRunning in user mode\n\n");
                    Requests.UserMode();
                    break;
                default:
                    Console.WriteLine("\nSelection not recognized. Running in user mode\n\n");
                    break;
            }

            // #region Setup Microsoft Graph Client

            //*********************************************************************
            // setup Microsoft Graph Client
            //*********************************************************************
            //Microsoft.Graph.GraphService client;
            //try
            //{
            //    client = AuthenticationHelper.GetActiveDirectoryClientAsUser();
            //}
            //catch (Exception ex)
            //{
            //    Console.ForegroundColor = ConsoleColor.Red;
            //    Console.WriteLine("Acquiring a token failed with the following error: {0}", ex.Message);
            //    if (ex.InnerException != null)
            //    {
            //        //You should implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
            //        //InnerException Message will contain the HTTP error status codes mentioned in the link above
            //        Console.WriteLine("Error detail: {0}", ex.InnerException.Message);
            //    }
            //    Console.ResetColor();
            //    Console.ReadKey();
            //    return;
            //}

            //#endregion


    //        try
    //        {
    //            #region Create a unified group
    //             // POST /groups - create unified group 
    //             // Uncomment once this is supported in PPE
    ////            string suffix = Helper.GetRandomString(5);
    ////            IList<string> grouplist = new List<string> {"Unified"};
    ////            Group group = new Group
    ////           {
    ////               groupTypes = grouplist,
    ////               displayName = "Unified group " + suffix,
    ////               description = "Group " + suffix + " is the best ever",
    ////               mailNickname = "Group" + suffix,
    ////               mailEnabled = true,
    ////               securityEnabled = false
    ////           };
    ////            try
    ////            {
    ////                client.groups.AddGroupAsync(group).Wait();
    ////                Console.WriteLine("\nCreated unified group {0}", group.displayName);
    ////            }
    ////            catch (Exception e)
    ////            {
    ////                Console.WriteLine("\nIssue creating the group {0}", group.displayName);
    ////                Console.WriteLine("\nUnexpected Error: {0} {1}", e.Message,
    ////e.InnerException != null ? e.InnerException.Message : "");
    ////            }
                
    //            #endregion


    //            #region Add group members
    //            //Group retrievedGroup = new Group();
    //            //List<IUser> members = client.users.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
    //            //List<IGroup> foundGroups = client.groups.Where(g => g.displayName.Equals(group.displayName)).ExecuteAsync().Result.CurrentPage.ToList();
    //            //if (foundGroups != null && foundGroups.Count > 0)
    //            //{
    //            //    retrievedGroup = foundGroups.First() as Group;
    //            //}
    //            //foreach (IUser _user in members)
    //            //{
    //            //    try
    //            //    {
    //            //        User member = (User)_user;
    //            //        group.members.Add(member as DirectoryObject);
                       
    //            //        //((IDirectoryObjectCollection)group.members).Add(member as DirectoryObject);
    //            //        group.UpdateAsync().Wait();
    //            //        Console.WriteLine("\nAdding {0} to group {1}", member.userPrincipalName, group.displayName);
    //            //    }
    //            //    catch (Exception e)
    //            //    {
    //            //        Console.WriteLine("\nError assigning member to group. {0} {1}",
    //            //             e.Message, e.InnerException != null ? e.InnerException.Message : "");
    //            //    }
    //            //}
                
                
    //            #endregion

            //#region Get groups
            //// GET /groups?$top=5
            //try
            //{
            //    List<Igroup> groups = client.groups.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
            //    Console.WriteLine();
            //    Console.WriteLine("GET /groups?$top=5");
            //    Console.WriteLine();
            //    foreach (Igroup _group in groups)
            //    {
            //        // Console.WriteLine("    Group Id: {0}  upn: {1} groupType: {2}", _group.objectId, _group.displayName, _group.groupType);
            //        Console.WriteLine("    Group Id: {0}  upn: {1}", _group.id, _group.displayName);
            //        foreach (string _type in _group.groupTypes)
            //        {
            //            if (_type == "Unified")
            //            {
            //                Console.WriteLine("       This is a Unifed Group");
            //            }
            //        }
            //    }
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine("\nFailed to read groups: {0} {1}", e.Message, e.InnerException != null ? e.InnerException.Message : "");
            //}
            //#endregion

    //            #region Get UNIFIED groups and their associated content
    //            // GET /groups?$top=5&$filter=groupType eq 'Unified'
    //            //groups = client.groups
    //            //    .Where(unifiedGroup => unifiedGroup.groupType.Equals("Unified"))
    //            //    .Take(5).ExecuteAsync().Result.CurrentPage.ToList();
    //            //Console.WriteLine();
    //            //Console.WriteLine("GET /groups?$top=5&$filter=groupType eq 'Unified'");
    //            //Console.WriteLine();
    //            //foreach (IGroup _group in groups)
    //            //{
    //            //    Console.WriteLine("    Unified Group: {0}", _group.displayName);

    //            //    try
    //            //    {
    //            //        List<IItem> unifiedGroupFiles = client.groups.GetById(_group.objectId).files.ExecuteAsync().Result.CurrentPage.ToList();
    //            //        if (unifiedGroupFiles.Count == 0)
    //            //        {
    //            //            Console.WriteLine("      no files for group");
    //            //        }
    //            //        foreach (IItem _file in unifiedGroupFiles)
    //            //        {
    //            //            Console.WriteLine("        file: {0} ", _file.name);
    //            //        }
    //            //        List<IDirectoryObject> unifiedGroupMembers = client.groups.GetById(_group.objectId).members.ExecuteAsync().Result.CurrentPage.ToList();
    //            //        if (unifiedGroupMembers.Count == 0)
    //            //        {
    //            //            Console.WriteLine("      no members for group");
    //            //        }
    //            //        foreach (IDirectoryObject _member in unifiedGroupMembers)
    //            //        {
    //            //            if (_member is IUser) {
    //            //                IUser memberUser = (IUser) _member;
    //            //                Console.WriteLine("        User: {0} ", memberUser.displayName);
    //            //            }
    //            //        }
    //            //        //TODO Add conversations, events and tasks when available
    //            //    }
    //            //    catch(Exception)
    //            //    {
    //            //        Console.Write("Unexpected exception when enumerating files for a group");
    //            //    }
    //            //}
    //            #endregion

    //            #region Get the signed in user's details, manager, reports and group memberships
    //            // GET /me
    //            try
    //            {

    //                User user = (User)client.Me.ExecuteAsync().Result;
    //                Console.WriteLine();
    //                Console.WriteLine("GET /me");
    //                Console.WriteLine();
    //                Console.WriteLine("    Id: {0}  UPN: {1}", user.objectId, user.userPrincipalName);
    //            }
    //            catch (Exception e)
    //            {
    //                Console.WriteLine("\nError getting files or content {0} {1}",
    //                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
    //            }


    //            //OAuth2Permission oauthperm = new OAuth2Permission
    //            //{
    //            //    id = new Guid("e61ff361-5baf-41f0-b2fd-380a6a5e406a"),
    //            //    adminConsentDescription = "foo",
    //            //    adminConsentDisplayName = "bar",
    //            //    isEnabled = true,
    //            //    type = "User",
    //            //    userConsentDescription = "user foo",
    //            //    userConsentDisplayName = "user bar",
    //            //    value = "my.scopename"
    //            //};

    //            //try
    //            //{
    //            //    Application app = new Application
    //            //    {
    //            //        displayName = "My api",
    //            //        identifierUris = new List<string> { "https://myapi.com" },
    //            //    };
    //            //    app.oauth2Permissions.Add(oauthperm);

    //            //    client.applications.AddApplicationAsync(app).Wait();

    //            //}
    //            //catch (Exception e)
    //            //{
    //            //    Console.WriteLine("Application Creation execption: {0} {1}", e.Message,
    //            //        e.InnerException != null ? e.InnerException.Message : "");
    //            //}

    //            //try
    //            //{
    //            //    User user = (User)client.users.GetById("sarad@adatumisv.onmicrosoft.com").ExecuteAsync().Result;
    //            //    Console.WriteLine();
    //            //    Console.WriteLine("GET /me");
    //            //    Console.WriteLine();
    //            //    Console.WriteLine("    Id: {0}  UPN: {1}", user.objectId, user.userPrincipalName);
    //            //}
    //            //catch (Exception e)
    //            //{
    //            //    Console.WriteLine("\nError getting files or content {0} {1}",
    //            //         e.Message, e.InnerException != null ? e.InnerException.Message : "");
    //            //}

               
    //            try
    //            {
    //                User user = (User)client.Me.ExecuteAsync().Result;
    //                String photo = user.thumbnailPhoto.SelfLink.AbsolutePath;
    //                string token = AuthenticationHelper.TokenForUser;
    //                Stream photoStream = Helper.GetRestRequestStream(photo, token).Result;
    //                Console.WriteLine("Got stream photo");
                    
    //               // IPhotoFetcher myphoto = (IPhotoFetcher)client.Me.UserPhoto.ToPhoto().ExecuteAsync().Result;
    //               // IPhoto _photo = myphoto.
    //            }
    //            catch (Exception)
    //            {
    //                Console.WriteLine("Failed to get stream");
    //            }

    //            // var foo = client.Me.files.addAsync("file.txt", "file2.text", "file", content);

    //            // GET /me/directReports
    //            List<IDirectoryObject> directs = client.Me.directReports.ExecuteAsync().Result.CurrentPage.ToList();
    //            Console.WriteLine();
    //            Console.WriteLine("GET /me/directReports");
    //            Console.WriteLine();
    //            if (directs.Count == 0)
    //            {
    //                Console.WriteLine("      no reports");
    //            }
    //            else
    //            {
    //                foreach (IDirectoryObject _user in directs)
    //                {
    //                    if (_user is IUser)
    //                    {
    //                        IUser __user = (IUser)_user;
    //                        Console.WriteLine("      Id: {0}  UPN: {1}", __user.objectId, __user.userPrincipalName);
    //                    }
    //                }
    //            }

    //            // GET /me/memberOf
    //            List<IDirectoryObject> _groups = client.Me.memberOf.ExecuteAsync().Result.CurrentPage.ToList();
    //            Console.WriteLine();
    //            Console.WriteLine("GET /me/memberOf");
    //            Console.WriteLine();
    //            if (_groups.Count == 0)
    //            {
    //                Console.WriteLine("    user is not a member of any groups");
    //            }
    //            foreach (IDirectoryObject _group in _groups)
    //            {
    //                if (_group is IGroup)
    //                {
    //                    IGroup __group= (IGroup)_group;
    //                    Console.WriteLine("    Id: {0}  UPN: {1}", __group.objectId, __group.displayName);
    //                }
    //            }
    //            #endregion

    //            #region Get signed in user's files, who last modified them, messages and events 
    //            // GET /me/drive/root/children?$top=5
    //            try
    //            {
    //                IList<Iitem> _items = client.Me.drive.root.children.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
    //                Console.WriteLine();
    //                Console.WriteLine("GET /me/files?$top=5");
    //                Console.WriteLine();
    //                foreach (Iitem _item in _items)
    //                {
                        
    //                    if (_item.file != null)
    //                    {
    //                        Console.WriteLine("    This is a folder: File Id: {0}  WebUrl: {1}", _item.id, _item.webUrl);
    //                    }
    //                    else {

    //                        Console.WriteLine("    File Id: {0}  WebUrl: {1}", _item.id, _item.webUrl);
    //                        DataServiceStreamLink link = _item.content;
    //                    }
    //                }
    //            }
    //            catch (Exception e)
    //            {
    //                Console.WriteLine("\nError getting files or content {0} {1}",
    //                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
    //            }

    //            // GET /me/Messages?$top=5
    //            List<IMessage> messages = client.Me.Messages.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
    //            Console.WriteLine();
    //            Console.WriteLine("GET /me/messages?$top=5");
    //            Console.WriteLine();
    //            if (messages.Count == 0)
    //            {
    //                Console.WriteLine("    no messages in mailbox");
    //            }
    //            foreach (IMessage message in messages)
    //            {
    //                Console.WriteLine("    Message: {0} received {1} ", message.Subject, message.DateTimeReceived);
    //            }

    //            // GET /me/Events?$top=5
    //            List<IEvent> events = client.Me.Events.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
    //            Console.WriteLine();
    //            Console.WriteLine("GET /me/events?$top=5");
    //            Console.WriteLine();
    //            if (events.Count == 0)
    //            {
    //                Console.WriteLine("    no events scheduled");
    //            }
    //            foreach (IEvent _event in events)
    //            {
    //                Console.WriteLine("    Event: {0} starts {1} ", _event.Subject, _event.Start);
    //            }
    //            #endregion

    //            #region Get the top 10 users, their files, and create a recipient list (to be used later)

    //            IList<Recipient> messageToList = new List<Recipient>();

    //            // GET /users?$top=5
    //            List<IUser> users = client.users.Take(10).ExecuteAsync().Result.CurrentPage.ToList();
    //            Console.WriteLine();
    //            Console.WriteLine("GET /users?$top=10 and their files");
    //            foreach (IUser _user in users)
    //            {
    //                Console.WriteLine(
    //                    "\n    User Id: {0}  upn: {1} license assigned: {2}", 
    //                    _user.objectId, 
    //                    _user.userPrincipalName, 
    //                    _user.assignedPlans != null && _user.assignedPlans.Count != 0 ? "true" : "false");
    //                if (_user.assignedPlans.Count != 0)
    //                {
    //                    Recipient messageTo = new Recipient();
    //                    EmailAddress emailAdress = new EmailAddress();
    //                    emailAdress.Address = _user.userPrincipalName;
    //                    emailAdress.Name = _user.displayName;
    //                    messageTo.EmailAddress = emailAdress;
    //                    messageToList.Add(messageTo);
    //                }

    //                //try
    //                //{
    //                //    if (_user.assignedPlans.Count != 0)
    //                //    {
    //                //        List<IItem> usersFiles = client.users.GetById(_user.objectId).files.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
    //                //        Console.WriteLine();
    //                //        Console.WriteLine("    GET /users/id/files?$top=5");
    //                //        if (usersFiles.Count != 0)
    //                //        {
    //                //            foreach (IItem _file in usersFiles)
    //                //            {
    //                //                Console.WriteLine("        file: {0} ", _file.name);
    //                //            }
    //                //        }
    //                //    }
    //                //}
    //                //catch(Exception)
    //                //{
    //                //    Console.WriteLine("\nUnexpected Error enumerating files of a user");
    //                //}
    //            }
    //            #endregion

    //            #region Send mail
    //            // POST /me/SendMail
    //            // TODO - debug why this keeps failing on setting ToRecipients..
    //            Console.WriteLine();
    //            Console.WriteLine("POST /me/sendmail");
    //            Console.WriteLine();

    //            try
    //            {
    //                ItemBody messageBody = new ItemBody();
    //                messageBody.Content = "Report pending";
    //                messageBody.ContentType = BodyType.Text;

    //                Message newMessage = new Message();
    //                newMessage.Subject = string.Format("\nCompleted test run from console app at {0}.", currentDateTime);
    //                newMessage.ToRecipients = (IList<Recipient>) messageToList;
    //                newMessage.Body = (ItemBody) messageBody;
                 
    //                client.Me.SendMailAsync(newMessage, true);
    //            }
    //            catch (Exception e)
    //            {
    //                Console.WriteLine("\nUnexpected Error attempting to send an email: {0} {1}", e.Message,
    //                e.InnerException != null ? e.InnerException.Message : "");
    //            }
    //            #endregion

    //            #region add a calendar event
    //            // POST /event
    //            try
    //            {
    //                Event newEvent = new Event();

    //                Attendee attendee = new Attendee();
    //                attendee.EmailAddress = new EmailAddress();
    //                attendee.EmailAddress.Name = "Alex Darrow";
    //                attendee.EmailAddress.Address = "alexd@imgeek.onmicrosoft.com";
    //                attendee.Type = AttendeeType.Required;
    //                newEvent.Attendees.Add(attendee);

    //                newEvent.Subject = "Discuss new project";
    //                newEvent.Start = DateTimeOffset.Parse("2015-05-05T18:00:00Z");
    //                newEvent.End = DateTimeOffset.Parse("2015-05-05T18:30:00Z");

    //                newEvent.Location = new Location();
    //                newEvent.Location.DisplayName = "Skype Call";

    //                newEvent.Body = new ItemBody();
    //                newEvent.Body.Content = "body";
    //                newEvent.Body.ContentType = BodyType.HTML;

    //                newEvent.Type = EventType.SingleInstance;

    //                client.Me.Events.AddEventAsync(newEvent).Wait();
    //            }
                

    //            catch (Exception e)
    //            {
    //                Console.WriteLine("\nUnexpected Error attempting to create a calendar event: {0} {1}", e.Message,
    //                e.InnerException != null ? e.InnerException.Message : "");
    //            }
    //            #endregion


    //        }
    //        catch (Exception e)
    //        {
    //            Console.WriteLine("\nUnexpected Error: {0} {1}", e.Message,
    //                e.InnerException != null ? e.InnerException.Message : "");
    //            throw;
    //        }


            //*********************************************************************************************
            // End of Demo Console App
            //*********************************************************************************************

            Console.WriteLine("\nCompleted at {0} \n Press Any Key to Exit.", currentDateTime);
            Console.ReadKey();
        }
    }
}
