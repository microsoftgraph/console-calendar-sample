using System.Threading.Tasks;
using NUnit.Framework;
using Calendar;
using Helpers;

namespace TestCalender
{
    [TestFixture]
    public class Tests
    {
        [Test]
        public async Task ScheduleEventTest()
        {
            var graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
            var calendar = new CalendarController(graphClient);

            var newEvent = await calendar
                .ScheduleEventAsync("A meeting", "16", "17", "meganb@M365x772687.onmicrosoft.com");

            Equals(newEvent.Subject, "A meeting");
        }
    }
}