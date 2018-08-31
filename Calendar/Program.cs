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

        static void Main(string[] args)
        {
            graphClient = Authentication.GetAuthenticatedClient();
            Console.WriteLine("Available commands: info, exit");
            var command = "";

            do
            {
                Console.Write("> ");
                command = Console.ReadLine();
                RunAsync(command).GetAwaiter().GetResult();
            }
            while (command != "exit");
        }

        static async Task RunAsync(string command)
        {
            switch (command)
            {
                case "info":
                    await GetMeAsync();
                    break;
                default:
                    Console.WriteLine("Command not available");
                    break;
            }
        }

        public static async Task GetMeAsync()
        {
            User user = null;
            try
            {
                user = await graphClient.Me.Request().GetAsync();

                Console.WriteLine($"Got user: {user.DisplayName}");
            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get the current user: " + e.Error.Message);
            }
        }
    }
}
