using System;
using System.Threading.Tasks;
using Helpers;
using Microsoft.Graph;
using System.Diagnostics;

namespace Calendar
{
    class Program
    {
        private static GraphServiceClient graphClient;
        private static CalendarController cal;

        static void Main(string[] args)
        {
            graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
            cal = new CalendarController(graphClient);
            RunAsync().GetAwaiter().GetResult();

            Console.WriteLine("Available commands:\n" +
                "\t 1. schedule \n " +
                "\t exit");
            var command = "";

            do
            {
                Console.Write("> ");
                command = Console.ReadLine();
                if (command != "exit") runAsync(command).GetAwaiter().GetResult();
            }
            while (command != "exit");
        }

        private static async Task runAsync(string command)
        {
            switch (command)
            {
                case "schedule":
                    Console.WriteLine("Enter the subject of your meeting");
                    var subject = Console.ReadLine();

                    await cal.ScheduleMeetingAsync(subject);
                    break;
                default:
                    Console.WriteLine("Invalid command");
                    break;
            }
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
            }
            else
            {
                Console.WriteLine("Did not find user");
            }

            Console.WriteLine();
        }
    }
}
