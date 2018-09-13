using Helpers;
using Microsoft.Graph;
using System;
using System.Diagnostics;
using System.Threading.Tasks;

namespace Calendar
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Welcome to the Calendar CLI");

            RunAsync().GetAwaiter().GetResult();
            Console.ReadKey();
        }
        /// <summary>
        /// Gets a User from Microsoft Graph
        /// </summary>
        /// <returns>A User object</returns>
        public static async Task<User> GetMeAsync()
        {
            User currentUser = null;
            try
            {
                var graphClient = GraphServiceClientProvider.GetAuthenticatedClient();

                // Request to get the current logged in user object from Microsoft Graph
                currentUser = await graphClient.Me.Request().GetAsync();

                return currentUser;
            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get the current user: " + e.Error.Message);
                return null;
            }
        }

        static async Task RunAsync()
        {
            var me = await GetMeAsync();

            if (me != null)
            {
                Console.WriteLine($"{me.DisplayName} logged in.");
            } else
            {
                Console.WriteLine("Did not find user");
            }

            Console.WriteLine();
        }
    }
}
