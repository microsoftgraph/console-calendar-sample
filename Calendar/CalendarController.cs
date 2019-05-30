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
        /// Schedules an event.
        /// 
        /// For purposes of simplicity we only allow the user to enter the startTime
        /// and endTime as hours.
        /// </summary>
        /// <param name="subject">Subject of the meeting</param>
        /// <param name="startTime">The time when the meeting starts</param>
        /// <param name="endTime">Duration of the meeting</param>
        /// <param name="attendeeEmail">Email of person to invite</param>
        /// <returns></returns>
        public async Task ScheduleEventAsync(string subject, string startTime, string endTime, string attendeeEmail)
        {
            DateTime dateTime = DateTime.Today;

            // set the start and end time for the event
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
        /// Sets recurrent events
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <returns></returns>
        public async Task SetRecurrentAsync(string subject, string startDate, string endDate, string startTime, string endTime)
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
            int startDay = int.Parse(startDate.Substring(0, 2));
            int startMonth = int.Parse(startDate.Substring(3, 2));
            int startYear = int.Parse(startDate.Substring(6, 4));

            int endDay = int.Parse(endDate.Substring(0, 2));
            int endMonth = int.Parse(endDate.Substring(3, 2));
            int endYear = int.Parse(endDate.Substring(6, 4));

            RecurrenceRange range = new RecurrenceRange
            {
                Type = RecurrenceRangeType.EndDate,
                StartDate = new Date(startYear, startMonth, startDay),
                EndDate = new Date(endYear, endMonth, endDay)
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

            DateTime dateTime = DateTime.Today;
            // set the start and end time for the event
            DateTimeTimeZone start = new DateTimeTimeZone
            {
                TimeZone = "Pacific Standard Time",
                DateTime = $"{startYear}-{startMonth}-{startDay}T{startTime}:00:00"
            };

            DateTimeTimeZone end = new DateTimeTimeZone
            {
                TimeZone = "Pacific Standard Time",
                DateTime = $"{startYear}-{startMonth}-{startDay}T{startTime}:00:00"
            };

            Event eventObj = new Event
            {
                Recurrence = recurrence,
                Subject = subject,
            };

            try
            {
                var recurrentEvent = await graphClient
                    .Me
                    .Events
                    .Request()
                    .AddAsync(eventObj);
                Console.WriteLine($"Created {recurrentEvent.Subject}," +
                    $" happens every week on Monday from {startTime}:00 to {endTime}:00");
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
        /// <param name="attendeeEmail"></param>
        /// <param name="date"></param>
        /// <returns></returns>
        public async Task SetAllDayAsync(string eventSubject, string attendeeEmail, string date)
        {
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

            int day = int.Parse(date.Substring(0, 2));
            int month = int.Parse(date.Substring(3, 2));
            int year = int.Parse(date.Substring(6, 4));

            Date allDayDate = new Date(year, month, day);
            DateTimeTimeZone start = new DateTimeTimeZone
            {
                TimeZone = "Pacific Standard Time",
                DateTime = allDayDate.ToString()
            };

            Date nextDay = new Date(year, month, day + 1);
            DateTimeTimeZone end = new DateTimeTimeZone
            {
                TimeZone = "Pacific Standard Time",
                DateTime = nextDay.ToString()
            };

            Event newEvent = new Event
            {
                Subject = eventSubject,
                Attendees = attendees,
                IsAllDay = true,
                Start = start,
                End = end
            };

            try
            {
                var allDayEvent = await graphClient
                    .Me
                    .Events
                    .Request()
                    .AddAsync(newEvent);

                Console.WriteLine($"Created an all day event: {newEvent.Subject}." +
                    $" Happening on {date}");
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

        public async Task<IUserEventsCollectionPage> GetEvents()
        {
            return await graphClient
                .Me
                .Events
                .Request()
                .GetAsync();
        }
    }
}
