using System;
using System.Linq;

namespace OutlookToGoogleCalenderPush
{
    internal class Program
    {
        public static void Log(string msg)
        {
            Console.Out.WriteLine($"{DateTime.Now:T} - {msg}");
        }

        private static void Main(string[] args)
        {
            EmbeddedLibsResolver.Init();

            if (args != null && args.Length == 1 && args[0].Equals("GetCalender", StringComparison.InvariantCultureIgnoreCase))
                GetCalender();

            if (args != null && args.Length == 2 && args[0].Equals("Sync", StringComparison.InvariantCultureIgnoreCase) && !String.IsNullOrEmpty(args[1]))
                Sync(args[1]);

            Console.Write("Press any key");
            Console.Read();
        }

        private static void Sync(string id)
        {
            Log($"Start {nameof(GoogleCalanderHelper.CreateCalendarService)}");
            var service = GoogleCalanderHelper.CreateCalendarService();

            Log($"Start {nameof(GoogleCalanderHelper.GetCalender)}");
            var cal = GoogleCalanderHelper.GetCalender(service, id);
            Log($"Found Calendar: Id: {cal.Id} - Summary: {cal.Summary}");

            Console.Out.WriteLine($"Are you sure? We will earse ALL events in this calander '{cal.Summary}'! [y/n]");
            var key = Console.ReadKey();
            if (key.Key != ConsoleKey.Y)
            {
                Console.Out.WriteLine($"This program will continue and ERASE all events in '{cal.Summary}'");
                return;
            }


            Log($"Start {nameof(OutlookHelper.GetAllCalendarItems)}");
            var result = OutlookHelper.GetAllCalendarItems();

            Log($"Get {result.Count} Events from Outlook.");

            if (!result.Any())
            {
                Log("Exit because nothing to do!");
                Console.ReadKey();
                return;
            }

            Log($"Start {nameof(GoogleCalanderHelper.DeleteGoogleCalenderEvents)}");
            GoogleCalanderHelper.DeleteGoogleCalenderEvents(service, id);

            Log($"Start {nameof(GoogleCalanderHelper.DeleteRecurringGoogleCalenderEvents)}");
            GoogleCalanderHelper.DeleteRecurringGoogleCalenderEvents(service, id);

            Log($"Start {nameof(GoogleCalanderHelper.AddEventToGoogleCalender)}");
            GoogleCalanderHelper.AddEventToGoogleCalender(service, result, id);
        }

        private static void GetCalender()
        {
            Log($"Start {nameof(GoogleCalanderHelper.CreateCalendarService)}");
            var service = GoogleCalanderHelper.CreateCalendarService();

            Log($"Start {nameof(GoogleCalanderHelper.GetCalanderList)}");
            var calanderList = GoogleCalanderHelper.GetCalanderList(service);

            Log("You need to pass the id to sync from outlook to google.");

            foreach (var entry in calanderList)
                Log($"    Id: {entry.Id} - Summary: {entry.Summary}");
        }
    }
}