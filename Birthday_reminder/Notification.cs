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
        public string getNotificationText(IEnumerable<KeyValuePair<string, DateTime>> birthdayList, NotificationFormat format)
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
