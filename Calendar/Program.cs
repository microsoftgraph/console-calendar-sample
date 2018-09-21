using Helpers;
using Microsoft.Graph;
using System;
using System.Diagnostics;
using System.Threading.Tasks;

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
                "\t 2. set-recurrent \n " +
                "\t 3. book-room \n " + 
                "\t 4. set-allday \n " +
                "\t 5. accept \n " +
                "\t 6. decline \n" +
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
                case "book-room":
                    Console.WriteLine("Enter the event id");
                    var eventId = Console.ReadLine();

                    Console.WriteLine("Enter the resource email");
                    var resourceEmail = Console.ReadLine();

                    await cal.BookRoomAsync(eventId, resourceEmail);
                    break;
                case "set-recurrent":
                    Console.WriteLine("Enter the event subject");
                    var eventSubject = Console.ReadLine();

                    await cal.SetRecurrentAsync(eventSubject);
                    break;
                case "set-allday":
                    Console.WriteLine("Enter the event's subject");
                    var allDaySubject = Console.ReadLine();

                    await cal.SetAllDayAsync(allDaySubject);
                    break;
                case "accept":
                    Console.WriteLine("Enter the event's id");
                    var eventToAccept = Console.ReadLine();

                    await cal.AcceptAsync(eventToAccept);
                    break;
                case "decline":
                    Console.WriteLine("Enter the event's id");
                    var eventToDecline = Console.ReadLine();

                    await cal.DeclineAsync(eventToDecline);
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
