#region

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Graph;
using System.Threading;

#endregion

namespace MicrosoftGraphSampleConsole
{
    internal class Requests
    {
        public static GraphServiceClient graphClient;

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
                graphClient = new GraphServiceClient(
                    new DelegateAuthenticationProvider(
                        (requestMessage) =>
                        {
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", AuthenticationHelper.GetTokenForUser());

                            return Task.FromResult(0);
                        }));
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

            var user = new User();
            try
            {

                user = graphClient.Me.Request().GetAsync().Result;
                Console.WriteLine();
                Console.WriteLine("GET /me");
                Console.WriteLine(); 
                Console.WriteLine("    Id: {0}  UPN: {1}", user.Id, user.UserPrincipalName);
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting /me user {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }

            try
            {
                // get signed in user's picture.  
                // Drop to REST for this - Client Library doesn't support this yet :(
                using (var photoStream = graphClient.Me.Photo.Content.Request().GetAsync().Result)
                {
                    Console.WriteLine("Got stream photo");
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Failed to get stream");
            }

            try
            {
                // GET /me/directReports
                var directsPage = graphClient.Me.DirectReports.Request().GetAsync().Result;
                Console.WriteLine();
                Console.WriteLine("GET /me/directReports");
                Console.WriteLine();
                if (directsPage.CurrentPage == null || directsPage.CurrentPage.Count == 0)
                {
                    Console.WriteLine("      no reports");
                }
                else
                {
                    foreach (var _user in directsPage.CurrentPage)
                    {
                        if (_user is User)
                        {
                            var __user = _user as User;
                            Console.WriteLine("      Id: {0}  UPN: {1}", __user.Id, __user.UserPrincipalName);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting directReports {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }

            try
            {
                // GET /me/manager
                var manager = graphClient.Me.Manager.Request().GetAsync().Result;
                Console.WriteLine();
                Console.WriteLine("GET /me/manager");
                Console.WriteLine();
                if (manager == null)
                {
                    Console.WriteLine("      no manager");
                }
                else
                {
                    var _user = graphClient.Users[manager.Id].Request().GetAsync().Result;
                    Console.WriteLine("\nManager      Id: {0}  UPN: {1}", _user.Id, _user.UserPrincipalName);
                    //    }
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
                var groupList = new List<DirectoryObject>();

                var currentPage = graphClient.Users[user.Id].MemberOf.Request().Top(1).GetAsync().Result;

                groupList.AddRange(currentPage);

                // Page through for the 5 results to test/demonstrate paging functionality as well
                for (int i=0; i < 4; i++)
                {
                    if (currentPage.NextPageRequest != null)
                    {
                        currentPage = currentPage.NextPageRequest.GetAsync().Result;

                        groupList.AddRange(currentPage);
                    }
                }

                Console.WriteLine();
                Console.WriteLine("GET /me/memberOf");
                Console.WriteLine();
                if (groupList.Count == 0)
                {
                    Console.WriteLine("    user is not a member of any groups");
                }

                foreach (var _group in groupList)
                {
                    if (_group is Group)
                    {
                        var __group = _group as Group;
                        Console.WriteLine("    Id: {0}  UPN: {1}", __group.Id, __group.DisplayName);
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
                var _items = graphClient.Me.Drive.Root.Children.Request().Top(5).GetAsync().Result;
                Console.WriteLine();
                Console.WriteLine("GET /me/files?$top=5");
                Console.WriteLine();
                foreach (var _item in _items)
                {

                    if (_item.File != null)
                    {
                        Console.WriteLine("    This is a folder: File Id: {0}  WebUrl: {1}", _item.Id, _item.WebUrl);
                    }
                    else
                    {

                        Console.WriteLine("    File Id: {0}  WebUrl: {1}", _item.Id, _item.WebUrl);
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
                var messages = graphClient.Me.Messages.Request().Top(5).GetAsync().Result;
                Console.WriteLine();
                Console.WriteLine("GET /me/messages?$top=5");
                Console.WriteLine();
                if (messages.CurrentPage == null || messages.Count == 0)
                {
                    Console.WriteLine("    no messages in mailbox");
                }
                foreach (var message in messages)
                {
                    Console.WriteLine("    Message: {0} received {1} ", message.Subject, message.ReceivedDateTime.Value);
                }

                // GET /me/Events?$top=5
                var events = graphClient.Me.Events.Request().Top(5).GetAsync().Result;
                Console.WriteLine();
                Console.WriteLine("GET /me/events?$top=5");
                Console.WriteLine();
                if (events.CurrentPage == null || events.Count == 0)
                {
                    Console.WriteLine("    no events scheduled");
                }
                foreach (var _event in events)
                {
                    Console.WriteLine("    Event: {0} starts {1} ", _event.Subject, _event.Start.DateTime);
                }

                // GET /me/contacts?$top=5
                var myContacts = graphClient.Me.Contacts.Request().Top(5).GetAsync().Result;
                Console.WriteLine();
                Console.WriteLine("GET /me/myContacts?$top=5");
                Console.WriteLine();
                if (myContacts.CurrentPage == null || myContacts.Count == 0)
                {
                    Console.WriteLine("    You don't have any contacts");
                }
                foreach (var _contact in myContacts)
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
            Console.WriteLine("\nSearch for user (enter search string):");
            String searchString = Console.ReadLine();
            
            IGraphServiceUsersCollectionPage searchResults = null;
            try
            {
                var queryOptions = new List<Option>
                {
                    new QueryOption("$filter",
                    string.Format(
                        "startswith(userPrincipalName,'{0}') or startswith(displayName,'{0}') or startswith(givenName,'{0}') or startswith(surname,'{0}')",
                        searchString)),
                };

                searchResults = graphClient
                    .Users
                    .Request(queryOptions)
                    .Top(5)
                    .GetAsync()
                    .Result;
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting User {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            if (searchResults.CurrentPage != null && searchResults.Count > 0)
            {
                foreach (var u in searchResults)
                {
                    Console.WriteLine("User {1} DisplayName: {0}",
                        u.DisplayName, u.UserPrincipalName);
                }
            }
            else
            {
                Console.WriteLine("User not found");
            }

            #endregion

            #region Create a unified group
            // POST /groups - create unified group 
            Console.WriteLine("\nDo you want to create a new unified group? Click y/n\n");
            var createGroup = Console.ReadLine();
            Group uGroup = null;

            try
            {
                if (string.Equals(createGroup, "y", StringComparison.OrdinalIgnoreCase) || string.Equals(createGroup, "yes", StringComparison.OrdinalIgnoreCase))
                {
                    string suffix = Helper.GetRandomString(5);
                    try
                    {
                        uGroup = graphClient.Groups.Request().AddAsync(
                            new Group
                            {
                                GroupTypes = new List<string> { "Unified" },
                                DisplayName = "Unified group " + suffix,
                                Description = "Group " + suffix + " is the best ever",
                                MailNickname = "Group" + suffix,
                                MailEnabled = true,
                                SecurityEnabled = false
                            }).Result;

                        Console.WriteLine("\nCreated unified group {0}", uGroup.DisplayName);
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("\nIssue creating the group Unified group {0}", suffix);
                        uGroup = null;
                    }
                }

                #endregion

                #region Add/remove group members

                Group retreivedGroup = uGroup;
                // get a set of users to add
                var members = graphClient.Users.Request().Top(3).GetAsync().Result;

                // Either add to newly created group, OR add to an existing group
                if (retreivedGroup == null)
                {
                    var foundGroups = graphClient
                        .Groups
                        .Request(new List<Option> { new QueryOption("$filter", "groupTypes/any(gt:gt%20eq%20'Unified')") })
                        .Top(5)
                        .GetAsync()
                        .Result;

                    if (foundGroups.CurrentPage != null && foundGroups.Count > 0)
                    {
                        retreivedGroup = foundGroups.First() as Group;
                    }
                }

                if (retreivedGroup != null)
                {
                    // Add users
                    foreach (var _user in members)
                    {
                        try
                        {
                            graphClient.Groups[retreivedGroup.Id].Members.References.Request().AddAsync(new User { Id = _user.Id }).Wait();

                            Console.WriteLine("\nAdded {0} to group {1}", _user.UserPrincipalName, retreivedGroup.DisplayName);
                        }
                        catch (AggregateException e)
                        {
                            e.Handle(exception =>
                            {
                                var serviceException = exception as ServiceException;

                                string errorDetail = null;

                                if (serviceException != null)
                                {
                                    errorDetail = serviceException.ToString();
                                }
                                else
                                {
                                    errorDetail = exception.Message;
                                }

                                Console.Write("\nError assigning member to group.\n{0}", errorDetail);

                                return true;
                            });
                        }
                    }

                    // now remove the added users
                    foreach (var _user in members)
                    {
                        try
                        {
                            graphClient.Groups[retreivedGroup.Id].Members[_user.Id].Reference.Request().DeleteAsync().Wait();

                            Console.WriteLine("\nRemoved {0} from group {1}", _user.UserPrincipalName, retreivedGroup.DisplayName);
                        }
                        catch (AggregateException e)
                        {
                            e.Handle(exception =>
                            {
                                var serviceException = exception as ServiceException;

                                string errorDetail = null;

                                if (serviceException != null)
                                {
                                    errorDetail = serviceException.ToString();
                                }
                                else
                                {
                                    errorDetail = exception.Message;
                                }

                                Console.Write("\nError removing member from group.\n{0}", errorDetail);

                                return true;
                            });
                        }
                    }
                }
                else
                {
                    Console.WriteLine("\nCan't find any unified groups to add members to.\n");
                }

                #endregion

                #region Get groups
                // GET /groups?$top=5
                var groups = graphClient.Groups.Request().Top(5).GetAsync().Result;
                Console.WriteLine();
                Console.WriteLine("GET /groups?$top=5");
                Console.WriteLine();
                foreach (var _group in groups)
                {
                    Console.WriteLine("    Group Id: {0}  upn: {1}", _group.Id, _group.DisplayName);
                    foreach (string _type in _group.GroupTypes)
                    {
                        if (_type == "Unified")
                        {
                            Console.WriteLine(": This is a Unifed Group");
                        }
                    }
                }
                #endregion

                #region Get the first 3 UNIFIED groups and view their associated content
                // GET /groups?$top=5&$filter=groupType eq 'Unified'
                groups = graphClient
                    .Groups
                    .Request(new List<Option> { new QueryOption("$filter", "groupTypes/any(gt:gt%20eq%20'Unified')") })
                    .Top(3)
                    .GetAsync()
                    .Result;
                Console.WriteLine();
                Console.WriteLine("GET /groups?$top=5&$filter=groupType eq 'Unified'");
                Console.WriteLine();
                foreach (var _group in groups)
                {
                    Console.WriteLine("    Unified Group: {0}", _group.DisplayName);

                    try
                    {
                        // get group members
                        var unifiedGroupMembers = graphClient.Groups[_group.Id].Members.Request().GetAsync().Result;
                        if (unifiedGroupMembers.CurrentPage == null || unifiedGroupMembers.Count == 0)
                        {
                            Console.WriteLine("      no members for group");
                        }
                        foreach (var _member in unifiedGroupMembers)
                        {
                            var _user = _member as User;
                            if (_user != null)
                            {
                                Console.WriteLine("        User: {0} ", _user.DisplayName);
                            }
                        }

                        //get group files
                        try
                        {
                            var unifiedGroupFiles = graphClient.Groups[_group.Id].Drive.Root.Children.Request().Top(5).GetAsync().Result;
                            if (unifiedGroupFiles.CurrentPage == null || unifiedGroupFiles.Count == 0)
                            {
                                Console.WriteLine("      no files for group");
                            }
                            foreach (var _file in unifiedGroupFiles)
                            {
                                Console.WriteLine("        file: {0} url: {1}", _file.Name, _file.WebUrl);
                            }
                        }
                        catch (Exception)
                        {
                            Console.Write("Unexpected exception when enumerating group files");
                        }

                        //get group conversations
                        try
                        {
                            var unifiedGroupConversations = graphClient.Groups[_group.Id].Conversations.Request().GetAsync().Result;
                            if (unifiedGroupConversations.CurrentPage == null || unifiedGroupConversations.Count == 0)
                            {
                                Console.WriteLine("      no conversations for group");
                            }
                            foreach (var _conversation in unifiedGroupConversations)
                            {
                                Console.WriteLine("        conversation topic: {0} ", _conversation.Topic);
                            }
                        }
                        catch (Exception)
                        {
                            Console.Write("Unexpected exception when enumerating group conversations");
                        }

                        //get group events
                        try
                        {
                            var unifiedGroupEvents = graphClient.Groups[_group.Id].Events.Request().GetAsync().Result;
                            if (unifiedGroupEvents.CurrentPage == null || unifiedGroupEvents.Count == 0)
                            {
                                Console.WriteLine("      no meeting events for group");
                            }
                            foreach (var _event in unifiedGroupEvents)
                            {
                                Console.WriteLine("        meeting event subject: {0} ", _event.Subject);
                            }
                        }
                        catch (Exception)
                        {
                            Console.Write("Unexpected exception when enumerating group events");
                        }
                    }
                    catch (Exception)
                    {
                        Console.Write("Unexpected exception when enumerating group members and events");
                    }
                }
                #endregion

                #region Get the top 10 users and create a recipient list (to be used later)
                var messageToList = new List<Recipient>();

                // GET /users?$top=5
                // Assigned plans aren't returned by default and *,assignedPlans select semantic doesn't work. Select all the
                // properties that we care about.
                var users = graphClient.Users.Request().Top(10).Select("userPrincipalName,displayName,emailAddress,assignedPlans").GetAsync().Result;
                foreach (var _user in users)
                {
                    if (_user.AssignedPlans != null && _user.AssignedPlans.Count() != 0)
                    {
                        var emailAdress = new EmailAddress
                        {
                            Address = _user.UserPrincipalName,
                            Name = _user.DisplayName,
                        };

                        var messageTo = new Recipient
                        {
                            EmailAddress = emailAdress,
                        };

                        messageToList.Add(messageTo);
                    }
                }

                // also add current signed in user to the recipient list IF they have a license
                if (user.AssignedPlans != null && user.AssignedPlans.Count() != 0)
                {
                    var emailAdress = new EmailAddress
                    {
                        Address = user.UserPrincipalName,
                        Name = user.DisplayName,
                    };

                    var messageTo = new Recipient
                    {
                        EmailAddress = emailAdress,
                    };

                    messageToList.Add(messageTo);
                }
                #endregion

                #region Send mail to signed in user and the recipient list
                // POST /me/SendMail

                Console.WriteLine();
                Console.WriteLine("POST /me/sendmail");
                Console.WriteLine();

                var messageBody = new ItemBody
                {
                    Content = "<report pending>",
                    ContentType = BodyType.Text,
                };

                var newMessage = new Message
                {
                    Subject = string.Format("\nCompleted test run from console app at {0}.", currentDateTime),
                    Body = messageBody,
                    ToRecipients = messageToList,
                };

                try
                {
                    graphClient.Me.SendMail(newMessage, true).Request().PostAsync().Wait();

                    Console.WriteLine("\nMail sent from {0}", user.DisplayName);
                }
                catch (AggregateException e)
                {
                    e.Handle(exception =>
                    {
                        var serviceException = exception as ServiceException;

                        string errorDetail = null;

                        if (serviceException != null)
                        {
                            errorDetail = serviceException.ToString();
                        }
                        else
                        {
                            errorDetail = exception.Message;
                        }

                        Console.Write("\nUnexpected Error attempting to send an email:\n{0}", errorDetail);

                        return true;
                    });
                }

                #endregion

                #region OneDrive functionality

                try
                {
                    var rootItems = graphClient.Me.Drive.Root.Children.Request().Top(5).GetAsync().Result;

                    if (rootItems.CurrentPage == null || rootItems.Count == 0)
                    {
                        Console.WriteLine("      no items found.");
                    }
                    foreach (var driveItem in rootItems)
                    {
                        Console.WriteLine("        item name: {0}     resource ID: {1}", driveItem.Name, driveItem.Id);
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine("\nUnexpected error attempting to get root items.");
                }

                DriveItem folder = null;

                try
                {
                    folder = graphClient.Me.Drive.Root.Children.Request().AddAsync(
                        new DriveItem
                        {
                            Folder = new Folder(),
                            Name = string.Format("Folder {0}", Helper.GetRandomString(5))
                        }).Result;

                    Console.WriteLine("\nCreated folder {0}", folder.Name);
                }
                catch (Exception)
                {
                    Console.WriteLine("\nUnexpected Error attempting to create folder.");
                }

                if (folder != null)
                {
                    try
                    {
                        var link = graphClient.Me.Drive.Items[folder.Id].CreateLink("view").Request().PostAsync().Result;

                        Console.WriteLine("\nCreated link {0}", link.Id);
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("\nUnexpected Error attempting to create link for folder.");
                    }

                    try
                    {
                        graphClient.Me.Drive.Items[folder.Id].Request().DeleteAsync().Wait();

                        Console.WriteLine("\nDeleted folder {0}", folder.Name);
                    }
                    catch (AggregateException e)
                    {
                        e.Handle(exception =>
                        {
                            Console.WriteLine("\nUnexpected Error attempting to delete folder.");

                            return true;
                        });
                    }
                }

                Console.WriteLine("\nSearch for drive items (enter search string):");
                var driveItemsSearchString = Console.ReadLine();

                try
                {
                    var driveSearchResults = graphClient.Me.Drive.Root.Search(driveItemsSearchString).Request().Top(10).GetAsync().Result;

                    if (driveSearchResults.CurrentPage == null || driveSearchResults.Count == 0)
                    {
                        Console.WriteLine("      no items found.");
                    }
                    foreach (var driveItem in driveSearchResults)
                    {
                        Console.WriteLine("        item name: {0}     resource ID: {1}", driveItem.Name, driveItem.Id);
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine("\nUnexpected Error attempting to create folder.");
                }

                #endregion

            }
            finally
            {
                #region clean up (delete any created items)
                if (uGroup != null)
                {
                    try
                    {
                        graphClient.Groups[uGroup.Id].Request().DeleteAsync().Wait();

                        Console.WriteLine("\nDeleted group {0}", uGroup.DisplayName);
                    }
                    catch (AggregateException e)
                    {
                        e.Handle(exception =>
                        {
                            var serviceException = exception as ServiceException;

                            string errorDetail = null;

                            if (serviceException != null)
                            {
                                errorDetail = serviceException.ToString();
                            }
                            else
                            {
                                errorDetail = exception.Message;
                            }

                            Console.Write("Couldn't delete group.  Error detail:\n{0}", errorDetail);

                            return true;
                        });
                    }
                }
                #endregion

                if (graphClient != null)
                {
                    graphClient.Dispose();
                }
            }

        }
        public static void AppMode()
        {
            // record start DateTime of execution
            string currentDateTime = DateTime.Now.ToUniversalTime().ToString();
            GraphServiceClient graphClient;
            #region Setup Microsoft Graph Client for app
            //*********************************************************************
            // setup Microsoft Graph Client for app
            //*********************************************************************
            try
            {
                graphClient = new GraphServiceClient(
                    new DelegateAuthenticationProvider(
                        (requestMessage) =>
                        {
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", AuthenticationHelper.GetTokenForApplication());

                            return Task.FromResult(0);
                        }));
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

            #region get specific user's email messages
            // 
            // get a specific user's email - demonstrates $select in Linq query
            try
            {
                Console.WriteLine("\nEnter the email address of the user mailbox you want to retrieve:");
                String email = Console.ReadLine();
                var messages = graphClient.Users[email].Messages.Request().Top(3).GetAsync().Result;
                Console.WriteLine();
                Console.WriteLine("GET /user/{0}/messages?$top=3&$select=subject,receivedDateTime", email);
                Console.WriteLine();
                if (messages.Count == 0)
                {
                    Console.WriteLine("    no messages in mailbox");
                }
                foreach (var message in messages)
                {
                    Console.WriteLine("    Message: {0} received {1} ", message.Subject, message.ReceivedDateTime);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting files or content {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }
            finally
            {
                if (graphClient != null)
                {
                    graphClient.Dispose();
                }
            }
            #endregion
        }
    }
}