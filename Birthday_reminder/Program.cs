using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Birthday_reminder
{
    public class Program
    {
        static void Main(string[] args)
        {
            #region Arranging arrays
            DateTime[] birthdays = new DateTime[0];
            string[] firstNames = new string[0];
            string[] secondNames = new string[0];
            double[] iterationNumers = new double[0];
            double[] iterationNumersforWeek = new double[0];

            #endregion

            #region WorkingWithExcel

            Application app = new Application();
            //app.Visible = true;

            Workbook pivotTableWorkbook = app.Workbooks.Open(@"C:\Users\yokvachuk\Desktop\Files\Enternainwork\Arbeit\41. Birthday_reminder\Birthday_reminder\BirthdayFileUpd.xlsx");
            Worksheet ws = pivotTableWorkbook.Worksheets["Birthdays"];
            #endregion

            #region RelevantArraysFetching

            string[] excelBirthdays = GetCellsDateTimeAsArray(ws, "C2", "C8");
            for (int x = 0; x < excelBirthdays.Length; x++)
            {
                DateTime[] tempArray = new DateTime[birthdays.Length + 1];
                for (int i = 0; i < birthdays.Length; i++)
                {
                    tempArray[i] = birthdays[i];
                }
                tempArray[birthdays.Length] = DateTime.Parse(excelBirthdays[x]); // length is by default a one more than iteration element number
                birthdays = tempArray;
            }

            string[] excelFirstName = GetCellsFirstNameAsArray(ws, "A2", "A8");
            for (int x = 0; x < excelFirstName.Length; x++)
            {
                string[] tempArray = new string[firstNames.Length + 1];
                for (int i = 0; i < firstNames.Length; i++)
                {
                    tempArray[i] = firstNames[i];
                }
                tempArray[firstNames.Length] = excelFirstName[x]; // length is by default a one more than iteration element number
                firstNames = tempArray;
            }

            string[] excelSecondName = GetCellsFirstNameAsArray(ws, "B2", "B8");
            for (int x = 0; x < excelSecondName.Length; x++)
            {
                string[] tempArray = new string[secondNames.Length + 1];
                for (int i = 0; i < secondNames.Length; i++)
                {
                    tempArray[i] = secondNames[i];
                }
                tempArray[secondNames.Length] = excelSecondName[x]; // length is by default a one more than iteration element number
                secondNames = tempArray;
            }
            #endregion

            #region Array Representations

            //foreach (var item in firstNames)
            //{
            //    Console.WriteLine(item);
            //}

            //foreach (var item in secondNames)
            //{
            //    Console.WriteLine(item);
            //}

            //foreach (var item in birthdays)
            //{
            //    Console.WriteLine(item);
            //}
            #endregion


            #region MatchingTodayWithBirthdays
            for (int i = 0; i < birthdays.Length; i++)
            {
                int nowday = DateTime.Now.Day;
                int nowmonth = DateTime.Now.Month;

                if (nowday == birthdays[i].Day && nowmonth == birthdays[i].Month)
                {
                    double[] temporaryiterationNumbers = new double[iterationNumers.Length + 1];
                    for (int k = 0; k < iterationNumers.Length; k++)
                    {
                        temporaryiterationNumbers[k] = iterationNumers[k];
                    }
                    temporaryiterationNumbers[iterationNumers.Length] = i;
                    iterationNumers = temporaryiterationNumbers;
                }
            }

            #endregion

            #region MatchingBirthdaysForWeek
            for (int i = 0; i < birthdays.Length; i++)
            {
                int nowday = DateTime.Now.Day;
                int nowmonth = DateTime.Now.Month;

                DateTime nowenumday = DateTime.Now;

                if (nowday == birthdays[i].Day && nowmonth == birthdays[i].Month ||
                    (string)nowenumday.AddDays(1).ToString("MM-dd") == (string)birthdays[i].ToString("MM-dd") ||
                    (string)nowenumday.AddDays(2).ToString("MM-dd") == (string)birthdays[i].ToString("MM-dd") ||
                    (string)nowenumday.AddDays(3).ToString("MM-dd") == (string)birthdays[i].ToString("MM-dd") ||
                    (string)nowenumday.AddDays(4).ToString("MM-dd") == (string)birthdays[i].ToString("MM-dd") ||
                    (string)nowenumday.AddDays(5).ToString("MM-dd") == (string)birthdays[i].ToString("MM-dd") ||
                    (string)nowenumday.AddDays(6).ToString("MM-dd") == (string)birthdays[i].ToString("MM-dd")
                    )
                {
                    double[] temporaryiterationNumbers = new double[iterationNumersforWeek.Length + 1];
                    for (int k = 0; k < iterationNumersforWeek.Length; k++)
                    {
                        temporaryiterationNumbers[k] = iterationNumersforWeek[k];
                    }
                    temporaryiterationNumbers[iterationNumersforWeek.Length] = i;
                    iterationNumersforWeek = temporaryiterationNumbers;
                }
            }

            #endregion


            #region ConsoleRepresentationForDay
            if (iterationNumers.Length == 1)
            {
                Console.WriteLine("Today there is the birthday of: {0} {1}", firstNames[(int)iterationNumers[0]], secondNames[(int)iterationNumers[0]]);
            }
            else
            {
                Console.WriteLine("Today there are birthdays of:");
                Console.WriteLine(" ");
                foreach (int item in iterationNumers)
                {
                    Console.WriteLine("{0} {1}", firstNames[item], secondNames[item]);
                }
            }
            #endregion

            #region ConsoleREpresentationForWeek

            Console.WriteLine("This week there are birthdays of:");
            Console.WriteLine(" ");

            foreach (int item in iterationNumersforWeek)
            {
                Console.WriteLine("{0} {1}", firstNames[item], secondNames[item]);
            }

            #endregion

        }
            #region MethodsforArrayextract


        public static string[] GetCellsDateTimeAsArray(Worksheet ws, string startCell, string endCell)
        {
            if (startCell == endCell)
            {
                return new string[] { "" + ws.Range[startCell].Value };
            }
            return ((Array)ws.Range[startCell + ":" + endCell].Cells.Value).OfType<object>().Select(o => o.ToString()).ToArray();
        }

        public static string[] GetCellsFirstNameAsArray(Worksheet ws, string startCell, string endCell)
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