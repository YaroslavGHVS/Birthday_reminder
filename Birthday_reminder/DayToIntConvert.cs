using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Birthday_reminder
{
    public class DayToIntConvert
    {
        public int getDayNumberFromUSAEnum(DayOfWeek dow) // method to transfer values from week enumerator to european order number
        {
            switch (dow)
            {
                case DayOfWeek.Monday:
                    return (int)0;
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
    }
}
