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
        /// Sets recurrent meetings
        /// </summary>
        /// <param name="eventId"></param>
        /// <returns></returns>
        public async Task SetRecurrentAsync(string eventId)
        {
            Event eventObj = new Event();

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

            eventObj.Recurrence = recurrence;

            try
            {            
                await graphClient
                    .Me
                    .Events[eventId]
                    .Request()
                    .UpdateAsync(eventObj);
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
            }
        }
    }
}
