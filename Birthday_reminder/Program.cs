using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Birthday_reminder
{
    public class Program
    {

        enum NotificationFormat
        {
            Day,
            Week
        }

        static void Main(string[] args)
        {
            BirthdaysList bdList = new BirthdaysList();

            DateTime currentDateNormalized = new DateTime(1, DateTime.Now.Month, DateTime.Now.Day);

            DayOfWeek currentDayOfWeek = DateTime.Now.DayOfWeek;
            Int32 currentDayNumber = getDayNumberFromUSAEnum(currentDayOfWeek);


            IEnumerable<KeyValuePair<string, DateTime>> currenWeekBirthdays = bdList.Dictionary.
                Where(x => x.Value >= currentDateNormalized.AddDays(-currentDayNumber)
                && x.Value <= currentDateNormalized.AddDays(6 - currentDayNumber));


            IEnumerable<KeyValuePair<string, DateTime>> currenDayBirthdays = bdList.Dictionary.Where(x => x.Value == currentDateNormalized);

            Console.WriteLine(getNotificationText(currenDayBirthdays, NotificationFormat.Day));
            Console.WriteLine(getNotificationText(currenWeekBirthdays, NotificationFormat.Week));
            Console.ReadKey();

            /*EmailClient.GmailClient mailClient = new EmailClient.GmailClient("a@gmail.com", "", false);
            mailClient.SendMail("TESTTTT", "voinn.andrey@gmail.com");*/

        }

        private static int getDayNumberFromUSAEnum(DayOfWeek dow) // method to transfer values from week enumerator to european order number
        {
            switch (dow)
            {
                case DayOfWeek.Monday:
                    return 0;
                case DayOfWeek.Tuesday:
                    return 1;
                case DayOfWeek.Wednesday:
                    return 2;
                case DayOfWeek.Thursday:
                    return 3;
                case DayOfWeek.Friday:
                    return 4;
                case DayOfWeek.Saturday:
                    return 5;
                case DayOfWeek.Sunday:
                    return 6;
                default:
                    throw new ArgumentException("Incorrect input argument on DateOfWeekConversion method. Can be (0-6). Current input: " + dow);
            }
        }

        private static string getNotificationText(IEnumerable<KeyValuePair<string, DateTime>> birthdayList, NotificationFormat format)
        {
            string outputText = "";

            if (format == NotificationFormat.Day)
            {
                switch (birthdayList.Count())
                {
                    case 0:
                        outputText += "";
                        break;
                    case 1:
                        outputText += "This day there is the birthday of:";
                        break;
                    default:
                        outputText += "Today there are birthdays of:";
                        break;
                }
            }
            else if (format == NotificationFormat.Week)
            {
                switch (birthdayList.Count())
                {
                    case 0:
                        outputText += "";
                        break;
                    case 1:
                        outputText += "This week there is the birthday of:";
                        break;
                    default:
                        outputText += "This week there are birthdays of:";
                        break;
                }
            }

            if (birthdayList.Count() > 0)
            {
                outputText += "\n";
                foreach (KeyValuePair<string, DateTime> item in birthdayList)
                {
                    outputText += (string.Format("Name: {0}, On: {1}", item.Key, item.Value.ToString("dd/MM")));
                    outputText += "\n";
                }
            }


            return outputText;
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