using System.Collections.Generic;
using System.Security.Policy;
using MailCLient;


namespace Birthday_reminder
{
    public class Program
    {
        static void Main(string[] args)
        {
            BirthdaysList bdList = new BirthdaysList(); // remains here
            Birthday_Definer bd_def = new Birthday_Definer(bdList);
            List<Record> birthdayListModified = new List<Record>();
            birthdayListModified = bd_def.GetModifiedRecordList(); // birthdayList marked

            // для того чтобы составить список пользователей, кому нужно отправить письмо
            List<string> receivers = new List<string>();
            foreach (Record record in birthdayListModified)
            {
                if (record.IsUserNotified)
                {
                    receivers.Add(record.Email);
                }
            }

            // для того чтобы получить текстовки
            string htmlDayNotifText = new Notification().getNotificationText(birthdayListModified, NotificationFormat.Day);
            string htmlWeekNotifText = new Notification().getNotificationText(birthdayListModified, NotificationFormat.Week);

            //send mail
            GmailClient gmailClient = new GmailClient();

            foreach (var item in birthdayListModified)
            {
                if (item.IsTodayBirthday == true)
                {
                    gmailClient.Send(birthdayListModified, htmlDayNotifText);
                    break;
                }
            }

            foreach (var item in birthdayListModified)
            {
                if (item.IsWeekBirthday == true)
                {
                    gmailClient.Send(birthdayListModified, htmlWeekNotifText);
                    break;
                }
            }

            //GmailClient gmail = new GmailClient();
            //gmail.Send(receivers, weekBirthdays);
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

