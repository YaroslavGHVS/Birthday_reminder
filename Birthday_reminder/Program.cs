using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Birthday_reminder
{
    public class Program
    {
        static void Main(string[] args)
        {
            BirthdaysList bdList = new BirthdaysList();

            DateTime currentDateNormalized = new DateTime(1, DateTime.Now.Month, DateTime.Now.Day);

            DayOfWeek currentDayOfWeek = DateTime.Now.DayOfWeek;
            
            Int32 currentDayNumber = new DayToIntConvert().getDayNumberFromUSAEnum(currentDayOfWeek); // weak reference

            IEnumerable<KeyValuePair<string, DateTime>> currenWeekBirthdays = bdList.Dictionary.
                Where(x => x.Value >= currentDateNormalized.AddDays(-currentDayNumber)
                && x.Value <= currentDateNormalized.AddDays(6 - currentDayNumber));


            IEnumerable<KeyValuePair<string, DateTime>> currenDayBirthdays = bdList.Dictionary.Where(x => x.Value == currentDateNormalized);

            Console.WriteLine(new Notification().getNotificationText(currenDayBirthdays, NotificationFormat.Day));
            Console.WriteLine(new Notification().getNotificationText(currenWeekBirthdays, NotificationFormat.Week));
            Console.ReadKey();

            /*EmailClient.GmailClient mailClient = new EmailClient.GmailClient("a@gmail.com", "", false);
            mailClient.SendMail("TESTTTT", "voinn.andrey@gmail.com");*/
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