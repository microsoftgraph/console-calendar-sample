using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace Helpers
{
    class GraphServiceClientProvider
    {
        // The client ID is used by the application to uniquely identify itself to the authentication endpoint.
        private static string clientId = ConfigurationManager.AppSettings["clientId"].ToString();
        private static string[] scopes = {
            "https://graph.microsoft.com/User.Read"
        };

        private static PublicClientApplication identityClientApp = new PublicClientApplication(clientId);
        private static GraphServiceClient graphClient = null;

        // Get an access token for the given context and resourceId. An attempt is first made to acquire the token silently.
        // If that fails, then we try to acquire the token by prompting the user.
        public static GraphServiceClient GetAuthenticatedClient()
        {
            if (graphClient == null)
            {
                try
                {
                    graphClient = new GraphServiceClient(
                        "https://graph.microsoft.com/v1.0",
                        new DelegateAuthenticationProvider(
                                async (requestMessage) =>
                                {
                                    var token = await getTokenForUserAsync();
                                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                                }
                            ));
                    return graphClient;
                } 
                catch(Exception error)
                {
                    Debug.WriteLine($"Could not create a graph client {error.Message}");
                }
            }
            return graphClient;
        }

        /// <summary>
        /// Get token for User
        /// </summary>
        /// <returns>Token for User</returns>
        private static async Task<string> getTokenForUserAsync()
        {
            AuthenticationResult authResult = null;

            try
            {
                IEnumerable<IAccount> account = await identityClientApp.GetAccountsAsync();
                authResult = await identityClientApp.AcquireTokenSilentAsync(scopes, account as IAccount);
                return authResult.AccessToken;
            }
            catch(MsalUiRequiredException error)
            {
                // This means the AcquireTokenSilentAsync threw an exception. 
                // This prompts the user to log in with their account so that we can get the token.
                authResult = await identityClientApp.AcquireTokenAsync(scopes);
                return authResult.AccessToken;
            }
        }

    }
}
