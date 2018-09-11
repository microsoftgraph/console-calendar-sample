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
        public async Task<Event> ScheduleMeetingAsync(string subject)
        {
            Event scheduledEvent = new Event();

            try
            {
                Event newEvent = new Event
                {
                    Subject = subject
                };

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
            EmailAddress email = new EmailAddress
            {
                Address = resourceMail
            };
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

        public async Task<Event> SetRecurrentAsync(string eventId)
        {
            Event updatedEvent = null;

            Event eventObj = new Event();

            RecurrencePattern pattern = new RecurrencePattern
            {
                Type = RecurrencePatternType.Daily,
                Interval = 1
            };

            List<Microsoft.Graph.DayOfWeek> daysOfWeek = new List<Microsoft.Graph.DayOfWeek>();
            daysOfWeek.Add(Microsoft.Graph.DayOfWeek.Monday);
            pattern.DaysOfWeek = daysOfWeek;

            RecurrenceRange range = new RecurrenceRange
            {
                Type = RecurrenceRangeType.EndDate,
                StartDate = new Date(2018, 11, 6),
                EndDate = new Date(2018, 11, 8)
            };

            PatternedRecurrence recurrence = new PatternedRecurrence
            {
                Pattern = pattern,
                Range = range
            };

            eventObj.Recurrence = recurrence;

            try
            {
                
                updatedEvent = await graphClient
                    .Me
                    .Events[eventId]
                    .Request()
                    .UpdateAsync(eventObj);

                Console.WriteLine($">>> {updatedEvent}");
            }
            catch (Exception error)
            {
                throw error;
            }

            return updatedEvent;
        }

        public async Task<Event> SetAlldayAsync(string eventId)
        {
            Event updatedEvent = null;

            Event patchObj = new Event();

            DateTimeTimeZone start = new DateTimeTimeZone
            {
                TimeZone = "Pacific Standard Time",
                DateTime = new Date(2018, 9, 6).ToString()
            };

            DateTimeTimeZone end = new DateTimeTimeZone
            {
                TimeZone = "Pacific Standard Time",
                DateTime = new Date(2018, 9, 8).ToString()
            };

            patchObj.IsAllDay = true;
            patchObj.Start = start;
            patchObj.End = end;

            try
            {
                updatedEvent = await graphClient
                    .Me
                    .Events[eventId]
                    .Request()
                    .UpdateAsync(patchObj);
            }
            catch (Exception)
            {

                throw;
            }

            return updatedEvent;
        }

        public async Task AcceptAsync(string eventId)
        {
            try
            {
                await graphClient
                      .Me
                      .Events[eventId]
                      .Accept()
                      .Request()
                      .PostAsync();
                    
            }
            catch (Exception)
            {

                throw;
            }
        }

        public async Task DeclineAsync(string eventId)
        {
            try
            {
                await graphClient
                    .Me
                    .Events[eventId]
                    .Decline()
                    .Request()
                    .PostAsync();
            } 
            catch (Exception)
            {
                throw;
            }
        }
    }
}
