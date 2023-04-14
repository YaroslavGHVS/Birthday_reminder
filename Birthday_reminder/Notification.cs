using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Birthday_reminder
{
    public enum NotificationFormat
    {
        Day,
        Week
    }

    public class Notification
    {
        string staticTextOfWeekNotification = "Greetings! <br> 🎉Birthdays on this week are: ";
        string statisTextOfDayNotification = "Greetings! <br> 🎉Today's Birthdays are: ";

        public string getNotificationText(List<Record> recordList, NotificationFormat format)
        {
            string outputText = "";

            if (format == NotificationFormat.Day)
            {
                outputText += statisTextOfDayNotification;
                
                outputText += "<br>";

                foreach (Record record in recordList)
                {
                    if (record.IsTodayBirthday)
                    {
                        outputText += (string.Format("{0}", record.Name));
                        outputText += "<br>";
                    }
                }
            }
            else if (format == NotificationFormat.Week)
            {
                outputText += staticTextOfWeekNotification;

                outputText += "<br>";
                
                foreach (Record record in recordList)
                {
                    if (record.IsWeekBirthday)
                    {
                        outputText += (string.Format("{0} at {1}", record.Name, record.BirthdayDate.ToString("dd.MM")));
                        outputText += "<br>";
                    }
                    
                }
            }
            return outputText;
        }
    }
 }
