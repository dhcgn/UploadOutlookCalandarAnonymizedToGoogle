using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace OutlookToGoogleCalenderPush
{
    public static class GoogleCalanderHelper
    {
        private const string ApplicationName = "OutlookToGoogleCalenderPush";
        private static readonly string[] Scopes = { CalendarService.Scope.Calendar };
        private static readonly string Folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SOtoG");

        public static CalendarService CreateCalendarService()
        {
            if (!Directory.Exists(Folder))
                Directory.CreateDirectory(Folder);

            UserCredential credential;

            var assembly = Assembly.GetExecutingAssembly();
            var name = assembly.GetManifestResourceNames().Single(s => s.EndsWith("client_id.json"));

            using (var stream = assembly.GetManifestResourceStream(name))
            {
                var credPath = Path.Combine(Folder, "creds.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }


        public static void AddEventToGoogleCalender(CalendarService service, List<OutlookHelper.MyEvent> result, string calendarId)
        {
            foreach (var myEvent in result)
            {
                var temp = service.Events.Insert(new Event()
                {
                    Summary = myEvent.SubjectReduced,
                    Location = myEvent.Location,

                    Start = new EventDateTime
                    {
                        DateTime = myEvent.Start,
                        TimeZone = "Europe/Berlin"
                    },
                    End = new EventDateTime
                    {
                        DateTime = myEvent.End,
                        TimeZone = "Europe/Berlin"
                    },
                }, calendarId);
                var newEvent = temp.Execute();
                Console.Out.WriteLine($"New Event: {newEvent.Id}");
            }

        }

        public static void DeleteGoogleCalenderEvents(CalendarService service, string calendarId)
        {
            EventsResource.ListRequest request = service.Events.List(calendarId);
            request.TimeMin = DateTime.Now.Date;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 1000;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // List events.
            Events events = request.Execute();
            Console.WriteLine("Upcoming events:");
            if (events.Items != null && events.Items.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                    string when = eventItem.Start.DateTime.ToString();
                    if (String.IsNullOrEmpty(when))
                    {
                        when = eventItem.Start.Date;
                    }
                    Console.WriteLine($"{eventItem.Summary} ({when}) {eventItem.Start.TimeZone} - {eventItem.Id} - {eventItem.RecurringEventId}");

                    var delete = service.Events.Delete(calendarId, eventItem.Id);
                    delete.Execute();
                }
            }
            else
            {
                Console.WriteLine("No upcoming events found.");
            }
        }


        public static void DeleteRecurringGoogleCalenderEvents(CalendarService service, string calendarId)
        {
            EventsResource.ListRequest request = service.Events.List(calendarId);
            request.TimeMin = DateTime.Now.Date;
            request.ShowDeleted = false;
            request.SingleEvents = false;
            request.MaxResults = 1000;

            Events events = request.Execute();
            if (events.Items != null && events.Items.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                    string when = eventItem.Start.DateTime.ToString();
                    if (String.IsNullOrEmpty(when))
                    {
                        when = eventItem.Start.Date;
                    }
                    Console.WriteLine($"DELETE: {eventItem.Summary} ({when}) {eventItem.Start.TimeZone} - {eventItem.Id} - {eventItem.RecurringEventId}");

                    var delete = service.Events.Delete(calendarId, eventItem.Id);
                    delete.Execute();
                }
            }
            else
            {
                Console.WriteLine("No upcoming events found.");
            }
        }

        public static IList<CalendarListEntry> GetCalanderList(CalendarService service)
        {
            var listRequest = service.CalendarList.List();
            var list = listRequest.Execute();
            return list.Items;
        }

        public static Calendar GetCalender(CalendarService service, string id)
        {
            var listRequest = service.Calendars.Get(id);
            return listRequest.Execute();
        }
    }
}
