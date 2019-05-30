using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Threading.Tasks;
using System.Configuration;
using System.Net.Http.Headers;
using System.Diagnostics;
using Microsoft.Graph.Auth;

namespace Calendar
{
    public class GraphServiceClientProvider
    {
        // The Client ID is used by the application to uniquely identify itself to the v2.0 authentication endpoint.
        static string clientId = ConfigurationManager.AppSettings["clientId"].ToString();
        private static string[] scopes = {
            "User.Read",
            "Calendars.ReadWrite"
        };
              
        private static GraphServiceClient graphClient = null;

        // Get an access token for the given context and resourceId. An attempt is first made to
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        public static GraphServiceClient GetAuthenticatedClient()
        {
            if (graphClient == null)
            {
                // Create Microsoft Graph client.
                try
                {
                    IPublicClientApplication clientApplication = InteractiveAuthenticationProvider.CreateClientApplication(clientId);
                    InteractiveAuthenticationProvider authProvider = new InteractiveAuthenticationProvider(clientApplication, scopes);

                    graphClient = new GraphServiceClient(authProvider);
                    return graphClient;
                }

                catch (Exception ex)
                {
                    Debug.WriteLine("Could not create a graph client: " + ex.Message);
                }
            }

            return graphClient;
        }
    }
}
