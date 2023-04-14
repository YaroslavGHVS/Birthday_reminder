using Org.BouncyCastle.Asn1.Bsi;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Birthday_reminder
{
    public class Birthday_Definer
    {
        public List<Record> DayBirthdays;

        private DateTime currentDateNormalized;
        private int currentDayNumber;
        private BirthdaysList bdList;

        public Birthday_Definer(BirthdaysList bdList)
        {
            currentDayNumber = new DayToIntConvert().getDayNumberFromUSAEnum(DateTime.Now.DayOfWeek); // weak reference
            currentDateNormalized = new DateTime(1, DateTime.Now.Month, DateTime.Now.Day);
            this.bdList = bdList;
        }

        public List<Record> GetModifiedRecordList()
        {
            foreach (var record in bdList.RecordList)
            {
                if (record.BirthdayDate == currentDateNormalized)
                {
                    //DayBirthdays.Add(bd);
                    record.IsTodayBirthday = true;
                    record.IsUserNotified = false;
                }

                if(record.BirthdayDate > currentDateNormalized && record.BirthdayDate <= currentDateNormalized.AddDays(7))
                {
                    record.IsWeekBirthday = true;
                    record.IsUserNotified = false;
                }
            }
            return bdList.RecordList;
        }
    }
}