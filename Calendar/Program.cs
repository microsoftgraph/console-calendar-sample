using System;
using System.Threading.Tasks;
using Helpers;
using Microsoft.Graph;
using System.Diagnostics;

namespace Calendar
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            RunAsync().GetAwaiter().GetResult();
            Console.ReadKey();
        }

        public static async Task<User> GetMeAsync()
        {
            User currentUserObject = null;
            try
            {
                var graphClient = Authentication.GetAuthenticatedClient();
                currentUserObject = await graphClient.Me.Request().GetAsync();

                Debug.WriteLine("Got user: " + currentUserObject.DisplayName);
                return currentUserObject;
            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get the current user: " + e.Error.Message);
                return null;
            }
        }

        static async Task RunAsync()
        {
            //Display information about the current user
            Console.WriteLine("Get My Profile");
            Console.WriteLine();

            var me = await GetMeAsync();

            Console.WriteLine(me.DisplayName);
            Console.WriteLine("User:{0}\t\tEmail:{1}", me.DisplayName, me.Mail);
            Console.WriteLine();

            //Display information about people near me
            Console.WriteLine("Get People Near Me");
        }
    }
}
