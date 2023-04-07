using System.Collections.Generic;
using MailCLient;


namespace Birthday_reminder
{
    public class Program
    {

        static void Main(string[] args)
        {

            BirthdaysList bdList = new BirthdaysList(); // remains here

            //Console.WriteLine(new Notification().getNotificationText(new Birthday_Definer(bdList).GetBirthdaysAtCurrentDay(), NotificationFormat.Day));
            string weekBirthdays = new Notification().getNotificationText(new Birthday_Definer(bdList).GetBirthdaysAtCurrentWeek(), NotificationFormat.Week);


            List<string> receivers = new List<string>();
            receivers.Add("yokvachuk@kpmg.ua");
            receivers.Add("avoinalovych@kpmg.ua");

            GmailClient gmail = new GmailClient();
            gmail.Send(receivers, weekBirthdays);
            
        }
    }
}

// References to learning materials

// https://www.youtube.com/watch?v=_Hn4hbe1NxM
// https://www.youtube.com/watch?v=93n2f80bK2k&t=38s
// https://stackoverflow.com/questions/25833425/read-all-rows-and-columns-using-microsoft-office-interop-excel
// https://www.dotnetperls.com/sort-dictionary
// https://www.tutorialsteacher.com/articles/convert-string-to-enum-in-cshar
// https://www.tutorialsteacher.com/articles/convert-string-to-enum-in-csharp

