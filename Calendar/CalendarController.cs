using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace Calendar
{
    class CalendarController
    {
        GraphServiceClient graphClient;

        public CalendarController(GraphServiceClient client)
        {
            graphClient = client;
        }

        /// <summary>
        /// Schedules a meeting.
        /// </summary>
        /// <param name="subject">Subject of the meeting</param>
        /// <param name="address">Physical address of the meeting</param>
        /// <returns></returns>
        public async Task<Event> ScheduleMeetingAsync(string subject)
        {
            Event scheduledEvent = new Event();

            try
            {
                Event newEvent = new Event();
                newEvent.Subject = subject;

                scheduledEvent = await graphClient
                    .Me
                    .Events
                    .Request()
                    .AddAsync(newEvent);


                Console.WriteLine($"Added {scheduledEvent.Subject}");
            }
            catch (ServiceException error)
            {
                Console.WriteLine(error.Message);
            }

            return scheduledEvent;
        }

        public async Task<Event> BookRoomAsync(string eventId, string resourceMail)
        {
            Event updatedEvent = null;
            Attendee room = new Attendee();
            EmailAddress email = new EmailAddress();

            email.Address = resourceMail;
            room.Type = AttendeeType.Resource;
            room.EmailAddress = email;

            try
            {
                List<Attendee> attendees = new List<Attendee>();
                Event patchEvent = new Event();

                attendees.Add(room);
                patchEvent.Attendees = attendees;

                updatedEvent = await graphClient
                    .Me
                    .Events[eventId]
                    .Request()
                    .UpdateAsync(patchEvent);

                Console.WriteLine(updatedEvent.Attendees.Count());
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
            }
            return updatedEvent;
        }
    }
}
