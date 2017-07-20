using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;

namespace OutlookToGoogleCalenderPush
{
    public static class OutlookHelper
    {
        private static Items GetAppointmentsInRange(
            Folder folder, DateTime startTime, DateTime endTime)
        {
            var filter = "[Start] >= '"
                         + startTime.ToString("g")
                         + "' AND [End] <= '"
                         + endTime.ToString("g") + "'";
            Debug.WriteLine(filter);
            try
            {
                var calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);
                var restrictItems = calItems.Restrict(filter);
                if (restrictItems.Count > 0)
                    return restrictItems;
                return null;
            }
            catch
            {
                return null;
            }
        }

        public static List<MyEvent> GetAllCalendarItems()
        {
            var result = new List<MyEvent>();

            var app = new Application();

            var nameSpace = app.GetNamespace("MAPI");
            ;
            var calendarFolder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar) as Folder;
            calendarFolder.Items.IncludeRecurrences = true;

            var start = DateTime.Now.Date;
            var end = start.AddMonths(1);
            var appointmentItems = GetAppointmentsInRange(calendarFolder, start, end);
            if (appointmentItems == null) return result;

            foreach (AppointmentItem appointmentItem in appointmentItems)
                result.Add(new MyEvent
                {
                    Subject = appointmentItem.Subject,
                    Start = appointmentItem.Start,
                    End = appointmentItem.End,
                    Location = appointmentItem.Location
                });
            return result;
        }

        public class MyEvent
        {
            public DateTime Start { get; set; }
            public DateTime End { get; set; }
            public string Subject { get; set; }
            public string SubjectReduced => this.Reduce(this.Subject);
            public string Location { get; set; }

            private string Reduce(string subject)
            {
                var rgx = new Regex("[^a-zA-Z0-9 -]");
                subject = rgx.Replace(subject, string.Empty);

                return subject.Split(new[] {' '}, StringSplitOptions.RemoveEmptyEntries)
                              .Select(s => s.All(Char.IsUpper) ? s : s.First().ToString())
                              .Aggregate((c, c1) => c + c1);
            }
        }
    }
}