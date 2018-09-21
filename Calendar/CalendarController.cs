using Microsoft.Graph;
using System;
using System.Collections.Generic;
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
        /// 
        /// For purposes of simplicity we only allow the user to enter the starTime
        /// and endTime as hours.
        /// </summary>
        /// <param name="subject">Subject of the meeting</param>
        /// <param name="startTime">The time when the meeting starts</param>
        /// <param name="endTime">Duration of the meeting</param>
        /// <param name="attendeeEmail">Email of person to invite</param>
        /// <returns></returns>
        public async Task ScheduleMeetingAsync(string subject, string startTime, string endTime, string attendeeEmail)
        {
            DateTime dateTime = DateTime.Today;

            // set the start and end time for the meeting
            DateTimeTimeZone start = new DateTimeTimeZone
            {
                TimeZone = "Pacific Standard Time",
                DateTime = $"{dateTime.Year}-{dateTime.Month}-{dateTime.Day}T{startTime}:00:00"
            };
            DateTimeTimeZone end = new DateTimeTimeZone
            {
                TimeZone = "Pacific Standard Time",
                DateTime = $"{dateTime.Year}-{dateTime.Month}-{dateTime.Day}T{endTime}:00:00"
            };

            // Adds attendee to the event
            EmailAddress email = new EmailAddress
            {
                Address = attendeeEmail
            };

            Attendee attendee = new Attendee
            {
                EmailAddress = email,
                Type = AttendeeType.Required,
            };
            List<Attendee> attendees = new List<Attendee>();
            attendees.Add(attendee);

            Event newEvent = new Event
            {
                Subject = subject,
                Attendees = attendees,
                Start = start,
                End = end
            };

            try
            {
                /**
                 * This is the same as a post request 
                 * 
                 * POST: https://graph.microsoft.com/v1.0/me/events
                 * Request Body
                 * {
                 *      "subject": <event-subject>
                 *      "start": {
                            "dateTime": "<date-string>",
                            "timeZone": "Pacific Standard Time"
                          },
                         "end": {
                             "dateTime": "<date-string>",
                             "timeZone": "Pacific Standard Time"
                          },
                          "attendees": [{
                            emailAddress: {
                                address: attendeeEmail 
                            }
                            "type": "required"
                          }]
                            
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

        /// <summary>
        /// Sets recurrent meetings
        /// </summary>
        /// <param name="subject"></param>
        /// <returns></returns>
        public async Task SetRecurrentAsync(string subject)
        {
            // Sets the event to happen every week
            RecurrencePattern pattern = new RecurrencePattern
            {
                Type = RecurrencePatternType.Weekly,
                Interval = 1
            };

            /**
             * Sets the days of the week the event occurs.
             * 
             * For this sample it occurs every Monday
             ***/
            List<Microsoft.Graph.DayOfWeek> daysOfWeek = new List<Microsoft.Graph.DayOfWeek>();
            daysOfWeek.Add(Microsoft.Graph.DayOfWeek.Monday);
            pattern.DaysOfWeek = daysOfWeek;

            /**
             * Sets the duration of time the event will keep recurring.
             * 
             * In this case the event runs from Nov 6th to Nov 26th 2018.
             **/
            RecurrenceRange range = new RecurrenceRange
            {
                Type = RecurrenceRangeType.EndDate,
                StartDate = new Date(2018, 11, 6),
                EndDate = new Date(2018, 11, 26)
            };

            /**
             * This brings together the recurrence pattern and the range to define the
             * PatternedRecurrence property.
             **/
            PatternedRecurrence recurrence = new PatternedRecurrence
            {
                Pattern = pattern,
                Range = range
            };

            Event eventObj = new Event
            {
                Recurrence = recurrence,
                Subject = subject
            };

            try
            {
                await graphClient
                    .Me
                    .Events
                    .Request()
                    .AddAsync(eventObj);
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
            }
        }

        /// <summary>
        /// Sets all day events
        /// </summary>
        /// <param name="eventSubject"></param>
        /// <returns></returns>
        public async Task SetAllDayAsync(string eventSubject)
        {
            DateTimeTimeZone start = new DateTimeTimeZone
            {
                TimeZone = "Pacific Standard Time",
                DateTime = new Date(2018, 12, 6).ToString()
            };
            DateTimeTimeZone end = new DateTimeTimeZone
            {
                TimeZone = "Pacific Standard Time",
                DateTime = new Date(2018, 12, 8).ToString()
            };

            Event newEvent = new Event
            {
                Subject = eventSubject,
                IsAllDay = true,
                Start = start,
                End = end,
            };

            try
            {
                var allDayEvent = await graphClient
                    .Me
                    .Events
                    .Request()
                    .AddAsync(newEvent);

                Console.WriteLine($"Created {newEvent.Subject}");
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
            }
        }

        /// <summary>
        /// Accepts an event invite
        /// </summary>
        /// <param name="eventId"></param>
        /// <returns></returns>
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
                Console.WriteLine("Accepted the event invite");

            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
            }
        }

        /// <summary>
        /// Declines an invite to an event
        /// </summary>
        /// <param name="eventId"></param>
        /// <returns></returns>
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
                Console.WriteLine("Event declined");

            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
            }
        }
    }
}
