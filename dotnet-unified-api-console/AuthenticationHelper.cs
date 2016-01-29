using System;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Graph;
using Microsoft.Graph.WindowsForms;

namespace MicrosoftGraphSampleConsole
{
    public class AuthenticationHelper
    {
        public static string TokenForUser;

        /// <summary>
        /// Get Active Directory Client for Application.
        /// </summary>
        /// <returns>ActiveDirectoryClient for Application.</returns>
        public static Task<IGraphServiceClient> GetActiveDirectoryClientAsApplication()
        {
            return BusinessClientExtensions.GetAuthenticatedClientUsingAppOnlyAuthentication(
                Constants.ClientIdForAppAuthn,
                Constants.redirectUriForUserAuthn,
                Constants.ClientSecret,
                Constants.TenantId);
        }

        /// <summary>
        /// Async task to acquire token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static Task<string> AcquireTokenAsyncForUser()
        {
            return Task.FromResult(GetTokenForUser());
        }

        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static string GetTokenForUser()
        {
            if (TokenForUser == null)
            {
                var redirectUri = new Uri(Constants.redirectUriForUserAuthn);
                string authority = Constants.AuthString + "common";
                AuthenticationContext authenticationContext = new AuthenticationContext(authority, false);
                AuthenticationResult userAuthnResult = authenticationContext.AcquireToken(Constants.ResourceUrl,
                    Constants.ClientIdForUserAuthn, redirectUri, PromptBehavior.RefreshSession);
                TokenForUser = userAuthnResult.AccessToken;
                Console.WriteLine("\n Welcome " + userAuthnResult.UserInfo.GivenName + " " +
                                  userAuthnResult.UserInfo.FamilyName);
            }

            return TokenForUser;
        }

        /// <summary>
        /// Get Active Directory Client for User.
        /// </summary>
        /// <returns>ActiveDirectoryClient for User.</returns>
        public static Task<IGraphServiceClient> GetActiveDirectoryClientAsUser()
        {
            return BusinessClientExtensions.GetAuthenticatedClient(
                Constants.ClientIdForUserAuthn,
                Constants.redirectUriForUserAuthn);
        }
    }
}
