using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Security.Cryptography.X509Certificates;

namespace Birthday_reminder
{
    public class Program
    {
        static void Main(string[] args)
        {
            #region WorkingWithExcel

            Application app = new Application();

            Workbook pivotTableWorkbook = app.Workbooks.Open(@"C:\Users\yokvachuk\Desktop\Files\Enternainwork\Arbeit\41. Birthday_reminder\Birthday_reminder\BirthdayFile4.xlsx");
            Worksheet ws = pivotTableWorkbook.Worksheets["Birthdays"];


            int totalColumns = ws.UsedRange.Columns.Count;
            int totalRows = ws.UsedRange.Rows.Count + 1;
            #endregion

            #region ArraysInitializing

            DateTime[] birthdays = new DateTime[totalRows];
            DateTime[] BirthdaysYearReset = new DateTime[totalRows];

            #endregion

            #region RelevantArraysFetching

            string[] excelBirthdays = GetCells(ws, "B1", "B"+totalRows);
            for (int x = 0; x < excelBirthdays.Length; x++)
            {
                birthdays[x] = DateTime.Parse(excelBirthdays[x]);
            }

            for (int i = 0; i < birthdays.Length; i++)
            {
                DateTime currdate = birthdays[i];
                Int32 birthmonth = currdate.Month;
                Int32 birthday = currdate.Day;

                BirthdaysYearReset[i] = new DateTime(1, birthmonth, birthday);
            }

            string[] namesList = GetCells(ws, "A1", "A"+totalRows);

            app.Workbooks.Close();
            #endregion

            #region Instantiating,Populating,SortingDictionary

            Dictionary<string, DateTime> birthdaydictionary = new Dictionary<string, DateTime>();
            Dictionary<string, DateTime> SortedBirthdayDictionary = new Dictionary<string, DateTime>();
            Dictionary<string, DateTime> WeekBirthdaysDictionary = new Dictionary<string, DateTime>();
            Dictionary<string, DateTime> DayBirthdaysDictionary = new Dictionary<string, DateTime>();

            for (int i = 0; i < totalRows - 1; i++) // there is less enumeration amounts than the number of rows by one
            {
                birthdaydictionary.Add(namesList[i], BirthdaysYearReset[i]);
            }

            List<DateTime> namesListSorted = birthdaydictionary.Values.ToList();
            namesListSorted.Sort();

            foreach (KeyValuePair<string, DateTime> item in birthdaydictionary.OrderBy(Value => Value.Value))
            {
                SortedBirthdayDictionary[item.Key] = item.Value;
            }
            #endregion

            #region TodayDataFetch

            DateTime NowCurrent = DateTime.Now;
            Int32 month = NowCurrent.Month;
            Int32 day = NowCurrent.Day;

            DateTime NowCurrentNormalized = new DateTime(1, month, day);

            #endregion

            #region WeekBirthdaysExtraction

            DayOfWeek NowDayOfWeek = DateTime.Now.DayOfWeek;

            int DateOfWeekConversion (DayOfWeek dow) // method to transfer values from week to number
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
                        throw new ArgumentException("We can't enter anythin");
                }
            }
            Int32 n = DateOfWeekConversion(NowDayOfWeek);

            IEnumerable<KeyValuePair<string, DateTime>> currenWeekBirthdaysDictionary = SortedBirthdayDictionary.Where(x => x.Value >= NowCurrentNormalized.AddDays(-n) && x.Value <= NowCurrentNormalized.AddDays(6-n)); //!!!

            foreach (KeyValuePair<string, DateTime> item in currenWeekBirthdaysDictionary)
            {
                WeekBirthdaysDictionary.Add(item.Key, item.Value);
            }

            #endregion

            #region DayBirthdayExtraction

            IEnumerable<KeyValuePair<string, DateTime>> currenDayBirthdaysDictionary = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized); //!!!

            foreach (KeyValuePair<string, DateTime> item in currenDayBirthdaysDictionary)
            {
                DayBirthdaysDictionary.Add(item.Key, item.Value);
            }

            #endregion

            #region ConsoleRepresentationForDay

            switch (DayBirthdaysDictionary.Count)
            {
                case 0:
                    Console.WriteLine("No employee has birthday today:");
                    break;
                case 1:
                    Console.WriteLine("This day there is the birthday of:");
                    break;
                default:
                    Console.WriteLine("Today there are birthdays of:");
                    break;
            }

            foreach (KeyValuePair<string, DateTime> item in DayBirthdaysDictionary)
            {
                //Console.WriteLine(item);
                int currentNowYear = DateTime.Now.Year;

                int Month = item.Value.Month;
                int Day = item.Value.Day;
                DateTime UpdatedDate = new DateTime(currentNowYear,Month, Day);

                DayOfWeek DayofWeek = item.Value.DayOfWeek;
                Console.WriteLine("Name: {0}, On: {1}", item.Key, UpdatedDate.ToString("dd/MM"));
            }
            Console.WriteLine("");

            #endregion

            #region ConsoleRepresentationForWeek

            switch (WeekBirthdaysDictionary.Count)
            {
                case 0:
                    Console.WriteLine("No employee has birthday this week:");
                    break;
                case 1:
                    Console.WriteLine("This week there is the birthday of:");
                    break;
                default:
                    Console.WriteLine("This week there are birthdas of:");
                    break;
            }

            foreach (KeyValuePair<string, DateTime> item in WeekBirthdaysDictionary)
            {
                int currentNowYear = DateTime.Now.Year;

                int Month = item.Value.Month;
                int Day = item.Value.Day;
                DateTime UpdatedDate = new DateTime(currentNowYear, Month, Day);

                DayOfWeek DayofWeek = item.Value.DayOfWeek;
                Console.WriteLine("Name: {0}, On: {1}", item.Key, UpdatedDate.ToString("dd/MM"));
            }

            #endregion
        }
            #region MethodsforArrayextract
        public static string[] GetCells(Worksheet ws, string startCell, string endCell)
        {
            if (startCell == endCell)
            {
                return new string[] { "" + ws.Range[startCell].Value };
            }
            return ((Array)ws.Range[startCell + ":" + endCell].Cells.Value).OfType<object>().Select(o => o.ToString()).ToArray();
        }

        #endregion
    }
}

// References to learning materials

// https://www.youtube.com/watch?v=_Hn4hbe1NxM
// https://www.youtube.com/watch?v=93n2f80bK2k&t=38s
// https://stackoverflow.com/questions/25833425/read-all-rows-and-columns-using-microsoft-office-interop-excel
// https://www.dotnetperls.com/sort-dictionary
// https://www.tutorialsteacher.com/articles/convert-string-to-enum-in-cshar
// https://www.tutorialsteacher.com/articles/convert-string-to-enum-in-csharp