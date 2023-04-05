using System;
using System.Collections.Generic;
using System.Linq;

namespace Birthday_reminder
{
    public class Birthday_Definer
    {
        private IEnumerable<KeyValuePair<string, DateTime>> currentDayBirthdays;
        private IEnumerable<KeyValuePair<string, DateTime>> currenWeekBirthdays;

        private DateTime currentDateNormalized;

        private int currentDayNumber;
        private BirthdaysList bdList;

        public Birthday_Definer(BirthdaysList bdList)
        {
            currentDayNumber = new DayToIntConvert().getDayNumberFromUSAEnum(DateTime.Now.DayOfWeek); // weak reference
            this.bdList = bdList;

            currentDateNormalized = new DateTime(1, DateTime.Now.Month, DateTime.Now.Day);
        }

        public IEnumerable<KeyValuePair<string, DateTime>> GetBirthdaysAtCurrentDay()
        {
            return bdList.Dictionary.Where(x => x.Value == currentDateNormalized);
        }

        public IEnumerable<KeyValuePair<string, DateTime>> GetBirthdaysAtCurrentWeek()
        {
            return bdList.Dictionary.Where(x => x.Value >= currentDateNormalized.AddDays(-currentDayNumber) && x.Value <= currentDateNormalized.AddDays(6 - currentDayNumber));
        }
    }
}