# Microsoft Graph Training Module - Using Microsoft Graph .NET SDK to interact with Outlook Calendar
In this demo you will create a .NET console application from scratch using .NET Framework 4.7.2, the Microsoft Graph SDK, and the Microsoft Authentication Library (MSAL).

## Register the application 
 
1. Navigate to the [the Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) to register your app. Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**. 
 
2. Select **New registration**. On the **Register an application** page, set the values as follows. 
 
* Set **Name** to **ConsoleDemoCalendar**. 
* Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**. 
* Leave **Redirect URI** empty. 
* Choose **Register**. 
 
3. On the **ConsoleDemoCalendar** page, copy the values of both the **Application (client) ID** and the **Directory (tenant) ID**. Save these two values, since you will need them later. 
 
4. Select the **Add a Redirect URI** link. On the **Redirect URIs** page, locate the **Suggested Redirect URIs for public clients (mobile, desktop)** section. Select the URI that begins with `msal` **and** the **urn:ietf:wg:oauth:2.0:oob** URI. 
 
5. Open the sample solution in Visual Studio and then open the **Constants.cs** file. Change the **Tenant** string to the **Directory (tenant) ID** value you copied earlier. Change the **ClientIdForUserAuthn** string to the **Application (client) ID** value. 

## Create the project in Visual Studio 2017

1. In Visual Studio 2017, create a new **Console Application** project targeting .NET Framework 4.7.2.

    ![Screenshot of Visual Studio 2017 new project menu.](../../Images/04.png)

1. Select **Tools > NuGet Package Manager > Package Manager Console**. In the console window, run the following commands:

    ```powershell
    Install-Package "Microsoft.Graph"
    Install-Package "Microsoft.Identity.Client" -Version 2.7.1
    Install-Package "System.Configuration.ConfigurationManager"
    ```

1. Edit the **app.config** file, and immediately before the `/configuration` element, add the following element replacing the value with the **Application ID** provided by the Application Registration Portal:

    ```xml
    <appSettings>
        <add key="clientId" value="YOUR APPLICATION ID"/>
    </appSettings>
    ```

    >Note: Make sure to replace the value with the **Application ID** value provided from the Application Registration Portal.

## Add AuthenticationHelper.cs

1. Add a class to the project named **AuthenticationHelper.cs**. This class will be responsible for authenticating using the Microsoft Authentication Library (MSAL), which is the **Microsoft.Identity.Client** package that we installed.

1. Replace the `using` statement at the top of the file.

    ```csharp
    using Microsoft.Graph;
    using Microsoft.Identity.Client;
    using System;
    using System.Configuration;
    using System.Diagnostics;
    using System.Linq;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    ```

1. Replace the `class` declaration with the following:

    ```csharp
   public class AuthenticationHelper
    {
        // The Client ID is used by the application to uniquely identify itself to the v2.0 authentication endpoint.
        static string clientId = ConfigurationManager.AppSettings["clientId"].ToString();
        public static string[] Scopes = { "Calendars.ReadWrite" };

        public static PublicClientApplication IdentityClientApp = new PublicClientApplication(clientId);

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
                    graphClient = new GraphServiceClient(
                        "https://graph.microsoft.com/v1.0",
                        new DelegateAuthenticationProvider(
                            async (requestMessage) =>
                            {
                                var token = await GetTokenForUserAsync();
                                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                            }));
                    return graphClient;
                }

                catch (Exception ex)
                {
                    Debug.WriteLine("Could not create a graph client: " + ex.Message);
                }
            }

            return graphClient;
        }

        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> GetTokenForUserAsync()
        {
            AuthenticationResult authResult = null;
            try
            {
                IEnumerable<IAccount> accounts = await IdentityClientApp.GetAccountsAsync();
                IAccount firstAccount = accounts.FirstOrDefault();

                authResult = await IdentityClientApp.AcquireTokenSilentAsync(Scopes, firstAccount);                
                return authResult.AccessToken;
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilentAsync.
                //This indicates you need to call AcquireTokenAsync to acquire a token

                authResult = await IdentityClientApp.AcquireTokenAsync(Scopes);

                return authResult.AccessToken;
            }
        }

        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static void SignOut()
        {
            foreach (var user in IdentityClientApp.GetAccountsAsync().Result)
            {
                IdentityClientApp.RemoveAsync(user);
            }
            graphClient = null;
        }
    }
    ```

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
