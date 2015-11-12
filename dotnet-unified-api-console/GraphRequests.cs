#region

using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.OData.Client;
using Microsoft.OData.ProxyExtensions;
using System.Net.Http;

#endregion

namespace MicrosoftGraphSampleConsole
{
    internal class Requests
    {
        public static Microsoft.Graph.GraphService client;
        public static void UserMode()
        {
            // record start DateTime of execution
            string currentDateTime = DateTime.Now.ToUniversalTime().ToString();
            #region Setup Microsoft Graph Client for user

            //*********************************************************************
            // setup Microsoft Graph Client for user...
            //*********************************************************************
            try
            {
                client = AuthenticationHelper.GetActiveDirectoryClientAsUser();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Acquiring a token failed with the following error: {0}", ex.Message);
                if (ex.InnerException != null)
                {
                    //You should implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
                    //InnerException Message will contain the HTTP error status codes mentioned in the link above
                    Console.WriteLine("Error detail: {0}", ex.InnerException.Message);
                }
                Console.ResetColor();
                Console.ReadKey();
                return;
            }

            #endregion

            Console.WriteLine("\nStarting user-mode requests...");
            Console.WriteLine("\n=============================\n\n");

            #region Get the signed in user's details, manager, reports and group memberships
            //// GET /me

            user user = new user();
            try
            {
                
                user = (user)client.me.ExecuteAsync().Result;
                Console.WriteLine();
                Console.WriteLine("GET /me");
                Console.WriteLine();
                Console.WriteLine("    Id: {0}  UPN: {1}", user.id, user.userPrincipalName);
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting /me user {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }

            //try
            //{
            //    User user = (User)client.Me.ExecuteAsync().Result;
            //    String photo = user.thumbnailPhoto.SelfLink.AbsolutePath;
            //    string token = AuthenticationHelper.TokenForUser;
            //    Stream photoStream = Helper.GetRestRequestStream(photo, token).Result;
            //    Console.WriteLine("Got stream photo");

            //    // IPhotoFetcher myphoto = (IPhotoFetcher)client.Me.UserPhoto.ToPhoto().ExecuteAsync().Result;
            //    // IPhoto _photo = myphoto.
            //}
            //catch (Exception)
            //{
            //    Console.WriteLine("Failed to get stream");
            //}

            try
            {
                // GET /me/directReports
                List<IdirectoryObject> directs = client.me.directReports.ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /me/directReports");
                Console.WriteLine();
                if (directs.Count == 0)
                {
                    Console.WriteLine("      no reports");
                }
                else
                {
                    foreach (IdirectoryObject _user in directs)
                    {
                        if (_user is Iuser)
                        {
                            Iuser __user = (Iuser)_user;
                            Console.WriteLine("      Id: {0}  UPN: {1}", __user.id, __user.userPrincipalName);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting directReports {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }

            // GET /me/memberOf
            try
            {
                List<IdirectoryObject> _groups = client.me.memberOf.ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /me/memberOf");
                Console.WriteLine();
                if (_groups.Count == 0)
                {
                    Console.WriteLine("    user is not a member of any groups");
                }
                foreach (IdirectoryObject _group in _groups)
                {
                    if (_group is Igroup)
                    {
                        Igroup __group = (Igroup)_group;
                        Console.WriteLine("    Id: {0}  UPN: {1}", __group.id, __group.displayName);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting group memberships {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            #region Get the signed in user's files, who last modified them, messages and events, and personal contacts
            try
            {
                IList<IdriveItem> _items = client.me.drive.root.children.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /me/files?$top=5");
                Console.WriteLine();
                foreach (IdriveItem _item in _items)
                {

                    if (_item.file != null)
                    {
                        Console.WriteLine("    This is a folder: File Id: {0}  WebUrl: {1}", _item.id, _item.webUrl);
                    }
                    else
                    {

                        Console.WriteLine("    File Id: {0}  WebUrl: {1}", _item.id, _item.webUrl);
                    }
                }
            }

            catch (Exception e)
            {
                Console.WriteLine("\nError getting files or content {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }

            try
            {

                // GET /me/Messages?$top=5
                List<Imessage> messages = client.me.Messages.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /me/messages?$top=5");
                Console.WriteLine();
                if (messages.Count == 0)
                {
                    Console.WriteLine("    no messages in mailbox");
                }
                foreach (Imessage message in messages)
                {
                    Console.WriteLine("    Message: {0} received {1} ", message.Subject, message.ReceivedDateTime);
                }

                // GET /me/Events?$top=5
                List<Ievent> events = client.me.Events.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /me/events?$top=5");
                Console.WriteLine();
                if (events.Count == 0)
                {
                    Console.WriteLine("    no events scheduled");
                }
                foreach (Ievent _event in events)
                {
                    Console.WriteLine("    Event: {0} starts {1} ", _event.Subject, _event.Start);
                }

                // GET /me/contacts?$top=5
                List<Icontact> myContacts = client.me.Contacts.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /me/myContacts?$top=5");
                Console.WriteLine();
                if (myContacts.Count == 0)
                {
                    Console.WriteLine("    You don't have any contacts");
                }
                foreach (Icontact _contact in myContacts)
                {
                    Console.WriteLine("    Contact: {0} ", _contact.DisplayName);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting messages, events and contacts {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }
            #endregion

            #region People picker example
            //*********************************************************************
            // People picker
            // Search for a user using text string "Sample" match against userPrincipalName, displayName, giveName, surname
            // Change searchString to suite your needs
            //*********************************************************************
            Console.WriteLine("Search for user you want in UPN,DisplayName,First or Last Name:\n");
            String searchString = Console.ReadLine();

            List<Iuser> usersList = null;
            IPagedCollection<Iuser> searchResults = null;
            try
            {
                IuserCollection userCollection = client.users;
                searchResults = userCollection.Where(u =>
                    u.userPrincipalName.StartsWith(searchString) ||
                    u.displayName.StartsWith(searchString) ||
                    u.givenName.StartsWith(searchString) ||
                    u.surname.StartsWith(searchString)).Take(5).ExecuteAsync().Result;
                usersList = searchResults.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting User {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            if (usersList != null && usersList.Count > 0)
            {
                do
                {
                    int j = 1;
                    usersList = searchResults.CurrentPage.ToList();
                    foreach (Iuser u in usersList)
                    {
                        Console.WriteLine("User {2} DisplayName: {0}  UPN: {1}",
                            u.displayName, u.userPrincipalName, j);
                        j++;
                    }
                    searchResults = searchResults.GetNextPageAsync().Result;
                } while (searchResults != null);
            }
            else
            {
                Console.WriteLine("User not found");
            }

            #endregion

            //    #region Create a unified group
            //     // POST /groups - create unified group 
            //    Console.WriteLine("\nDo you want to create a new unified group? Click y/n\n");
            //    ConsoleKeyInfo key = Console.ReadKey();
            //    if (key.KeyChar == 'y')
            //    {
            //        string suffix = Helper.GetRandomString(5);
            //        group group = new group
            //        {
            //            groupTypes = new List<string> { "Unified" },
            //            displayName = "Unified group " + suffix,
            //            description = "Group " + suffix + " is the best ever",
            //            mailNickname = "Group" + suffix,
            //            mailEnabled = true,
            //            securityEnabled = false
            //        };
            //        try
            //        {
            //            client.groups.AddgroupAsync(group).Wait();
            //            Console.WriteLine("\nCreated unified group {0}", group.displayName);
            //        }
            //        catch (Exception)
            //        {
            //            Console.WriteLine("\nIssue creating the group {0}", group.displayName);
            //        }
            //    }                
            //    #endregion

            //    #region Add group members
            //    //client.Context.SaveChanges();
            //    //Group retrievedGroup = new Group(); 
            //    //List<IUser> members = client.users.Take(3).ExecuteAsync().Result.CurrentPage.ToList();
            //    ////List<IGroup> foundGroups = client.groups.Where(g => g.displayName.Equals(group.displayName)).ExecuteAsync().Result.CurrentPage.ToList();
            //    ////if (foundGroups != null && foundGroups.Count > 0)
            //    ////{
            //    ////    retrievedGroup = foundGroups.First() as Group;
            //    ////}
            //    //foreach (IUser _user in members)
            //    //{
            //    //    try
            //    //    {
            //    //        group.members.Add(_user as DirectoryObject);
            //    //        group.UpdateAsync().Wait();
            //    //        Console.WriteLine("\nAdding {0} to group {1}", _user.userPrincipalName, group.displayName);
            //    //    }
            //    //    catch (Exception e)
            //    //    {
            //    //        Console.WriteLine("\nError assigning member to group. {0} {1}",
            //    //             e.Message, e.InnerException != null ? e.InnerException.Message : "");
            //    //    }
            //    //}

            //    //// now remove the first one
            //    //foreach (IUser _user in members)
            //    //{
            //    //    try
            //    //    {
            //    //        User member = (User)_user;
            //    //        group.members.Remove(member as DirectoryObject);
            //    //        group.UpdateAsync().Wait();
            //    //        Console.WriteLine("\nRemoved {0} from group {1}", member.userPrincipalName, group.displayName);
            //    //        // only remove the first one
            //    //        break;
            //    //    }
            //    //    catch (Exception e)
            //    //    {
            //    //        Console.WriteLine("\nError removing member from group. {0} {1}",
            //    //             e.Message, e.InnerException != null ? e.InnerException.Message : "");
            //    //    }
            //    //}
                
                
            //    #endregion

                #region Get groups
                // GET /groups?$top=5
                List<Igroup> groups = client.groups.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /groups?$top=5");
                Console.WriteLine();
                foreach (Igroup _group in groups)
                {
                    Console.WriteLine("    Group Id: {0}  upn: {1}", _group.id, _group.displayName);
                    foreach (string _type in _group.groupTypes)
                    {
                        if (_type == "Unified")
                        {
                            Console.WriteLine(": This is a Unifed Group");
                        }
                    }
                }
                #endregion

                #region Get the top 5 UNIFIED groups and view their associated content
                // GET /groups?$top=5&$filter=groupType eq 'Unified'
                groups = client.groups.Where(ug => ug.groupTypes.Any(gt =>gt == "Unified")).Take(5).ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /groups?$top=5&$filter=groupType eq 'Unified'");
                Console.WriteLine();
                foreach (Igroup _group in groups)
                {
                    Console.WriteLine("    Unified Group: {0}", _group.displayName);

                    try
                    {
                        //get group members
                        List<IdirectoryObject> unifiedGroupMembers = client.groups.GetById(_group.id).members.ExecuteAsync().Result.CurrentPage.ToList();
                        if (unifiedGroupMembers.Count == 0)
                        {
                            Console.WriteLine("      no members for group");
                        }
                        foreach (IdirectoryObject _member in unifiedGroupMembers)
                        {
                            if (_member is Iuser)
                            {
                                Iuser memberUser = (Iuser)_member;
                                Console.WriteLine("        User: {0} ", memberUser.displayName);
                            }
                        }

                        //get group files
                        IList<IdriveItem> unifiedGroupFiles = client.groups.GetById(_group.id).drive.root.children.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
                        if (unifiedGroupFiles.Count == 0)
                        {
                            Console.WriteLine("      no files for group");
                        }
                        foreach (IdriveItem _file in unifiedGroupFiles)
                        {
                            Console.WriteLine("        file: {0} url: {1}", _file.name, _file.webUrl);
                        }    
                                        
                        //get group conversations
                        List<Iconversation> unifiedGroupConversations = client.groups.GetById(_group.id).Conversations.ExecuteAsync().Result.CurrentPage.ToList();
                        if(unifiedGroupConversations.Count == 0)
                        {                                                    
                            Console.WriteLine("      no conversations for group");
                        }
                        foreach (Iconversation _conversation in unifiedGroupConversations)
                        {
                            Console.WriteLine("        conversation topic: {0} ", _conversation.Topic);
                        }

                        //get group events
                        List<Ievent> unifiedGroupEvents = client.groups.GetById(_group.id).Events.ExecuteAsync().Result.CurrentPage.ToList();
                        if(unifiedGroupEvents.Count == 0)
                        {                                                    
                            Console.WriteLine("      no meeting events for group");
                        }
                        foreach (Ievent _event in unifiedGroupEvents)
                        {
                            Console.WriteLine("        meeting event subject: {0} ", _event.Subject);
                        }
                    }
                    catch (Exception)
                    {
                        Console.Write("Unexpected exception when enumerating group contents and members");
                    }
                }
                #endregion

                #region Get the top 5 users and create a recipient list (to be used later)
                IList<recipient> messageToList = new List<recipient>();

                // GET /users?$top=5
                List<Iuser> users = client.users.Take(10).ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /users?$top=10 and their files");
                foreach (Iuser _user in users)
                {
                    Console.WriteLine(
                        "\n    User Id: {0}  upn: {1} license assigned: {2}", 
                        _user.id, 
                        _user.userPrincipalName, 
                        _user.assignedPlans != null && _user.assignedPlans.Count != 0 ? "true" : "false");
                    if (_user.assignedPlans.Count != 0)
                    {
                        recipient messageTo = new recipient();
                        emailAddress emailAdress = new emailAddress();
                        emailAdress.Address = _user.userPrincipalName;
                        emailAdress.Name = _user.displayName;
                        messageTo.EmailAddress = emailAdress;
                        messageToList.Add(messageTo);
                    }
                }

                // also add current signed in user to the recipient list IF they have a license
                if (user.assignedPlans.Count != 0)
                {               
                    recipient messageTo = new recipient();
                    emailAddress emailAdress = new emailAddress();
                    emailAdress.Address = user.userPrincipalName;
                    emailAdress.Name = user.displayName;
                    messageTo.EmailAddress = emailAdress;
                    messageToList.Add(messageTo);
                }
                #endregion

                #region Send mail to signed in user and the recipient list
                // POST /me/SendMail
 
                Console.WriteLine();
                Console.WriteLine("POST /me/sendmail");
                Console.WriteLine();

                try
                {
                    itemBody messageBody = new itemBody();
                    messageBody.Content = "<report pending>";
                    messageBody.ContentType = BodyType.Text;

                    message newMessage = new message();
                    newMessage.Subject = string.Format("\nCompleted test run from console app at {0}.", currentDateTime);
                    newMessage.ToRecipients = (IList<recipient>)messageToList;
                    newMessage.Body = (itemBody)messageBody;

                    client.me.SendMailAsync(newMessage, true);

                    Console.WriteLine("\nMail sent to {0}", user.displayName);
                }
                catch (Exception)
                {
                    Console.WriteLine("\nUnexpected Error attempting to send an email");
                    throw;
                }
                #endregion

        }
        public static void AppMode()
        {
            // record start DateTime of execution
            string currentDateTime = DateTime.Now.ToUniversalTime().ToString();
            #region Setup Microsoft Graph Client for app
            //*********************************************************************
            // setup Microsoft Graph Client for app
            //*********************************************************************
            try
            {
                client = AuthenticationHelper.GetActiveDirectoryClientAsApplication();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Acquiring a token failed with the following error: {0}", ex.Message);
                if (ex.InnerException != null)
                {
                    //You should implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
                    //InnerException Message will contain the HTTP error status codes mentioned in the link above
                    Console.WriteLine("Error detail: {0}", ex.InnerException.Message);
                }
                Console.ResetColor();
                Console.ReadKey();
                return;
            }
            #endregion

            Console.WriteLine("\nStarting app-mode requests...");
            Console.WriteLine("\n=============================\n\n");
            try
            {
                string email = "sarad@adatumisv.onmicrosoft.com";
                List<Imessage> messages = client.users.GetById(email).Messages.Take(3).
                    ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /user/{0}/messages?$top=3", email);
                Console.WriteLine();
                if (messages.Count == 0)
                {
                    Console.WriteLine("    no messages in mailbox");
                }
                foreach (Imessage message in messages)
                {
                    Console.WriteLine("    Message: {0} received {1} ", message.Subject, message.ReceivedDateTime);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting files or content {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }
        }
    }
}