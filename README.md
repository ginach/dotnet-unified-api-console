# ASP.NET console sample using the Microsoft Graph client library

This sample shows how to connect and use Microsoft Graph (previously called Office 365 unified API) using the a pre-release of the Microsoft Graph client library. It uses [Active Directory Authentication Library](https://msdn.microsoft.com/en-us/library/azure/jj573266.aspx) to acquire OAuth2 tokens to call the Microsoft Graph.  This console app can operate in 2 modes: in the traditional delegated user mode (aka app+user mode), and/or in application mode (where no user sign in is required and the app accesses the API as itself).

## Prerequisites

To use this sample, you need the following:
* Visual Studio 2013 or 2015 installed and working on your development computer. 

* An Office 365 account. You can sign up for [an Office 365 Developer subscription](https://portal.office.com/Signup/Signup.aspx?OfferId=6881A1CB-F4EB-4db3-9F18-388898DAF510&DL=DEVELOPERPACK&ali=1#0) that includes the resources that you need to start building Office 365 apps.

     > Note: If you already have a subscription, the previous link sends you to a page with the message *Sorry, you can’t add that to your current account*. In that case use an account from your current Office 365 subscription.
* A Microsoft Azure tenant to register your application. Azure Active Directory (AD) provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

     > Important: You also need to make sure your Azure subscription is bound to your Office 365 tenant. To do this, see the Active Directory team's blog post, [Creating and Managing Multiple Windows Azure Active Directories](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx). The section **Adding a new directory** will explain how to do this. You can also see [Set up your Office 365 development environment](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) and the section **Associate your Office 365 account with Azure AD to create and manage apps** for more information.
* You'll need to register 2 applications in Azure:
  1. The client ID and redirect URI values of a native client application registered in Azure, for the delegated user mode. This application must be granted the NEED TO GET THE LIST OF PERMS **Send mail as signed-in user** permission for the **Microsoft Graph**. [Add a web application in Azure](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp) and [grant the proper permissions](https://github.com/OfficeDev/O365-AspNetMVC-Microsoft-Graph-Connect/wiki/Grant-permissions-to-the-Connect-application-in-Azure) to it.
  2. The client ID and secret of a web application also registered in Azure, for the application mode.
  3. You'll also need the tenant ID for the tenant you are registering these applications in.

## Configure and run the app
1. Open **NewName** file. 
2. In Solution Explorer, open the **Constants.cs** file.
3. Follow the instructions in the **Constants.cs** file to populate the properties for the 2 apps.

## Questions and comments

We'd love to get your feedback about the Office 365 365 ASP.NET MVC Connect sample. You can send your questions and suggestions to us in the [Issues](https://github.com/OfficeDev/WHATEVER/issues) section of this repository.

Questions about Microsoft Graph development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/MicrosoftGraph). Make sure that your questions or comments are tagged with [Office365] and [MicrosoftGraph].
  
## Additional resources

* [Microsoft Graph documentation](http://graph.microsoft.io)
* [Microsoft Graph API References](http://graph.microsoft.io/docs/api-reference/v1.0)


## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.