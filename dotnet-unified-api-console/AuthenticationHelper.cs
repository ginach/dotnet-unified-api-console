using System;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace MicrosoftGraphSampleConsole
{
    public class AuthenticationHelper
    {
        public static string TokenForUser;
        public static DateTimeOffset expiration;

        /// <summary>
        /// Get Token for Application.
        /// </summary>
        /// <returns>Token for application.</returns>
        public static string GetTokenForApplication()
        {
            string authority = Constants.AuthString + Constants.TenantId;
            AuthenticationContext authenticationContext = new AuthenticationContext(authority, false);
            // Config for OAuth client credentials 
            ClientCredential clientCred = new ClientCredential(Constants.ClientIdForAppAuthn, Constants.ClientSecret);
            AuthenticationResult authenticationResult = authenticationContext.AcquireToken(Constants.ResourceUrl,
                clientCred);
            string token = authenticationResult.AccessToken;
            return token;
        }

        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static string GetTokenForUser()
        {
            if (TokenForUser == null || expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
            {
                var redirectUri = new Uri(Constants.redirectUriForUserAuthn);
                string authority = Constants.AuthString + "common";
                AuthenticationContext authenticationContext = new AuthenticationContext(authority, false);
                AuthenticationResult userAuthnResult = authenticationContext.AcquireToken(Constants.ResourceUrl,
                    Constants.ClientIdForUserAuthn, redirectUri, PromptBehavior.RefreshSession);
                TokenForUser = userAuthnResult.AccessToken;
                expiration = userAuthnResult.ExpiresOn;
                Console.WriteLine("\n Welcome " + userAuthnResult.UserInfo.GivenName + " " +
                                  userAuthnResult.UserInfo.FamilyName);
            }

            return TokenForUser;
        }
    }
}
