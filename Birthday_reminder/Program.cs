using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
//using System.Linq;

namespace Birthday_reminder
{
        //public static class Extensions
        //{
        //    public static void Append<K,V>(this Dictionary<K,V> first, Dictionary<K,V> second)
        //    {
        //        List<KeyValuePair<K, V>> pairs = second.ToList();
        //        pairs.ForEach(pair => first.Add(pair.Key, pair.Value));
        //    }
            
        //}

    public class Program
    {

        static void Main(string[] args)
        {
            #region WorkingWithExcel

            Application app = new Application();

            Workbook pivotTableWorkbook = app.Workbooks.Open(@"C:\Users\yokvachuk\Desktop\Files\Enternainwork\Arbeit\41. Birthday_reminder\Birthday_reminder\BirthdayFile2.xlsx");
            Worksheet ws = pivotTableWorkbook.Worksheets["Birthdays"];

            int totalColumns = ws.UsedRange.Columns.Count;
            int totalRows = ws.UsedRange.Rows.Count;
            #endregion

            #region ArraysInitializing

            DateTime[] birthdays = new DateTime[totalRows];
            DateTime[] BirthdaysYearReset = new DateTime[totalRows];
            string[] namesList = new string[totalRows];

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
                var month = currdate.Month;
                var day = currdate.Day;

                BirthdaysYearReset[i] = new DateTime(1, month, day);
            }

            string[] excelNames = GetCells(ws, "A1", "A"+totalRows);
            for (int x = 0; x < excelNames.Length; x++)
            {
                namesList[x] = excelNames[x];
            }

            #endregion

            #region PopulatingAndFillingDictionary

            Dictionary<string, DateTime> birthdaydictionary = new Dictionary<string, DateTime>();
            Dictionary<string, DateTime> SortedBirthdayDictionary = new Dictionary<string, DateTime>();
            Dictionary<string, DateTime> WeekBirthdaysDictionary = new Dictionary<string, DateTime>();
            Dictionary<string, DateTime> DayBirthdaysDictionary = new Dictionary<string, DateTime>();

            for (int i = 0; i < totalRows - 1; i++) // there is less enumeration amounts than the number of rows by one
            {
                birthdaydictionary.Add(namesList[i], BirthdaysYearReset[i]); //!!!
            }
            
            var namesListSorted = birthdaydictionary.Values.ToList();
            namesListSorted.Sort();

            foreach (var item in birthdaydictionary.OrderBy(Value => Value.Value))
            {
                SortedBirthdayDictionary[item.Key] = item.Value;
            }

            #endregion

            var DayofWeekNow = DateTime.Now.DayOfWeek;

            var NowCurrent = DateTime.Now;
            var NowMonth = NowCurrent.Month;
            var NowDay = NowCurrent.Day;
            DateTime NowCurrentNormalized = new DateTime(1, NowMonth, NowDay);

            #region MatchingBirthdaysForWeek


            switch (DayofWeekNow)
            {
                case DayOfWeek.Monday:
                    
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized);

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                           WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(1)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(1));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(2)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(2));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(3)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(3));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(4)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(4));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(5)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(5));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(6)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(6));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    break;

                case DayOfWeek.Tuesday:
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-1)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-1));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized);

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(1)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(1));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(2)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(2));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(3)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(3));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(4)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(4));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(5)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(5));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    break;
                case DayOfWeek.Wednesday:
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-2)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-2));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-1)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-1));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized);

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(1)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(1));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(2)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(2));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(3)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(3));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(4)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(4));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    break;
                case DayOfWeek.Thursday:
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-3)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-3));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-2)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-2));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-1)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-1));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized);

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(1)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(1));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(2)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(2));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(3)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(3));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    break;
                case DayOfWeek.Friday:
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-4)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-4));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-3)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-3));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-2)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-2));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-1)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-1));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized);

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(1)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(1));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(2)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(2));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    break;
                case DayOfWeek.Saturday:
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-5)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-5));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-4)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-4));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-3)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-3));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-2)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-2));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-1)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-1));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized);

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(1)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(1));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    break;
                case DayOfWeek.Sunday:
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-6)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-6));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-5)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-5));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-4)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-4));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-3)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-3));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-2)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-2));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized.AddDays(-1)))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized.AddDays(-1));

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized))
                    {
                        var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized);

                        foreach (KeyValuePair<string, DateTime> item in currentData)
                        {
                            WeekBirthdaysDictionary.Add(item.Key, item.Value);
                        }
                    }
                    break;
                default:
                    Console.WriteLine("No birthdays this week.");
                    break;
            }

            #endregion

            #region MatchingBirthdayForDay

            if (SortedBirthdayDictionary.ContainsValue(NowCurrentNormalized))
            {
                var currentData = SortedBirthdayDictionary.Where(x => x.Value == NowCurrentNormalized);

                foreach (KeyValuePair<string, DateTime> item in currentData)
                {
                    DayBirthdaysDictionary.Add(item.Key, item.Value);
                }
            }

            #endregion


            #region ConsoleRepresentationForDay

            Console.WriteLine("This day there is the birthday of:");
            foreach (var item in DayBirthdaysDictionary)
            {
                Console.WriteLine(item);
            }
            Console.WriteLine("");

            #endregion

            #region ConsoleREpresentationForWeek

            Console.WriteLine("This week there are the birthdays of:");
            foreach (var item in WeekBirthdaysDictionary)
            {
                Console.WriteLine(item);
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
// @"C:\Users\yokvachuk\Desktop\Files\Enternainwork\Arbeit\41. Birthday_reminder\Birthday_reminder\BirthdayFile.xlsx"

// https://www.youtube.com/watch?v=_Hn4hbe1NxM
// https://www.youtube.com/watch?v=93n2f80bK2k&t=38s

// https://stackoverflow.com/questions/25833425/read-all-rows-and-columns-using-microsoft-office-interop-excel
// https://www.dotnetperls.com/sort-dictionary