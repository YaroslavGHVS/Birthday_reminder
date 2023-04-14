using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Birthday_reminder
{
    public class Record
    {
        public string Name;
        public DateTime BirthdayDate;
        public string Email;
        public bool IsUserNotified;
        public bool IsTodayBirthday;
        public bool IsWeekBirthday;

        public Record(string name, DateTime birthday, string email, bool IsTodayBirthday = false, bool IsWeekBIrthday = false, bool notification = true)
        {
            this.Name = name;
            this.BirthdayDate = birthday;
            this.Email = email;
            this.IsUserNotified = notification;
            this.IsTodayBirthday = IsTodayBirthday;
            this.IsWeekBirthday = IsWeekBIrthday;
        }
    }
}