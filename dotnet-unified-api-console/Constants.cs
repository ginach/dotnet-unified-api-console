namespace MicrosoftGraphSampleConsole
{
    internal class Constants
    {
        // To run this console app you need to have 2 application registrations - as the app runs in 2 modes...

        // app only configuration
        // In the Azure management portal, register a web application, and get a clientId and secret/key, 
        // and configure permissions according to README
        public const string TenantId = "08c33862-2dbb-4696-942d-6c31ca525938";
        public const string ClientIdForAppAuthn = "f33644b6-88f2-4b1e-a31d-8f0a126531f1";
        public const string ClientSecret = "lDkEtgj4pEisZZ7FApyG0oG+MFz9WVdU1ENLzk+PT0g=";

        // app+user console app configuration
        // In the Azure management portal, register a native client application, 
        // and configure permissions according to README
        public const string ClientIdForUserAuthn = "22990061-de54-42ef-b78a-a7ec69c1146e";
        public const string redirectUriForUserAuthn = "http://localhost:44323";

        public const string AuthString = "https://login.microsoftonline.com/";
        public const string ResourceUrl = "https://graph.microsoft.com/";
        public const string Url = "https://graph.microsoft.com/v1.0/";

    }
}