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
            graphClient = Authentication.GetAuthenticatedClient();
            cal = new CalendarController(graphClient);

            Console.WriteLine("Available commands: info, schedule, exit");
            var command = "";

            do
            {
                Console.Write("> ");
                command = Console.ReadLine();
                runAsync(command).GetAwaiter().GetResult();
            }
            while (command != "exit");
        }

        private static async Task runAsync(string command)
        {
            switch (command)
            {
                case "info":
                    await getMeAsync();
                    break;
                case "schedule":
                    Console.WriteLine("Enter the subject of your meeting");
                    var subject = Console.ReadLine();

                    await cal.ScheduleMeetingAsync(subject);
                    break;
                default:
                    Console.WriteLine("You've done it! You discovered Drake's Fortune.");
                    break;
            }
        }

        private static async Task getMeAsync()
        {
            try
            {
                User user = await graphClient.Me.Request().GetAsync();

                Console.WriteLine($"Got user: {user.DisplayName}");
            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get the current user: " + e.Error.Message);
            }
        }
    }
}
