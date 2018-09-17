using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

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
        public async Task ScheduleMeetingAsync(string subject)
        {
            Event newEvent = new Event();
            newEvent.Subject = subject;

            try
            {
                /**
                 * This is the same as a post request 
                 * 
                 * POST: https://graph.microsoft.com/v1.0/me/events
                 * Request Body
                 * {
                 *      "subject": <event-subject>
                 * }
                 * 
                 * Learn more about the properties of an Event object in the link below
                 * https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/event
                 * */
                Event calendarEvent = await graphClient
                    .Me
                    .Events
                    .Request()
                    .AddAsync(newEvent);

                Console.WriteLine($"Added {calendarEvent.Subject}");
            }
            catch (ServiceException error)
            {
                Console.WriteLine(error.Message);
            }

        }

        /// <summary>
        /// Books a room for the event
        /// </summary>
        /// <param name="eventId"></param>
        /// <param name="resourceMail"></param>
        /// <returns></returns>
        public async Task BookRoomAsync(string eventId, string resourceMail)
        {
            /**
             * A room is an an attendee of type resource
             * 
             * Refer to the link below to learn more about the properties of the Attendee class
             * https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/attendee
             **/
            Attendee room = new Attendee();
            EmailAddress email = new EmailAddress();
            email.Address = resourceMail;
            room.Type = AttendeeType.Resource;
            room.EmailAddress = email;

            List<Attendee> attendees = new List<Attendee>();
            Event patchEvent = new Event();

            attendees.Add(room);
            patchEvent.Attendees = attendees;

            try
            {
                /**
                 * This is the same as making a patch request
                 * 
                 * PATCH https://graph.microsoft.com/v1.0/me/events/{id}
                 * 
                 * request body 
                 * {
                 *      attendees: [{
                 *              emailAddress: {
                 *                  "address": "email@address.com"
                 *              },
                 *              type: "resource"
                 *          }
                 *      ]
                 * }
                 * */
                 await graphClient
                    .Me
                    .Events[eventId]
                    .Request()
                    .UpdateAsync(patchEvent);
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
            }
        }
    }
}
