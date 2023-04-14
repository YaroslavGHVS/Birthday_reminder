using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;


namespace Birthday_reminder
{
    /// <summary>
    /// Класс, который подготавливает все вводные для работы с таблицей дней рождений, представленной в excel.
    /// Фильтрация не входит в функционал класса и выполняется за его пределами
    /// </summary>
    public class BirthdaysList
    {
        //public Dictionary<string, DateTime> Dictionary;
        public List<Record> RecordList;

        public int ColumnsAmount { get; private set; }
        public int RowsAmount { get; private set; }

        private List<DateTime> dates;
        private List<string> names;
        private List<string> emails;

        private Worksheet worksheet;
        private ExcelImporter importer;

        public BirthdaysList()
        {
            RecordList = new List<Record>();

            importer = new ExcelImporter();
            worksheet = importer.GetFirstWorksheet();

            ColumnsAmount = worksheet.UsedRange.Columns.Count;
            RowsAmount = worksheet.UsedRange.Rows.Count + 1;

            dates = new List<DateTime>();
            
            fillObjectTable();
        }
        
        /// <summary>
        /// Метод инкапсулирует служебные методы для наполнения объекта
        /// </summary>
        private void fillObjectTable()
        {
            names = getCellsRange(worksheet, "A1", "A" + RowsAmount);

            List<string> excelBirthdays = getCellsRange(worksheet, "B1", "B" + RowsAmount);
            foreach (string birthday in excelBirthdays)
            {
                DateTime date = DateTime.Parse(birthday);
                dates.Add(new DateTime(1, date.Month, date.Day));
            }

            emails = getCellsRange(worksheet, "C1", "C" +  RowsAmount);

            fillNameDateDictionary();
        }

        private void fillNameDateDictionary()
        {
            for (int i = 0; i < RowsAmount - 1; i++)
                RecordList.Add(new Record(names[i], dates[i], emails[i]));
        }

        private List<string> getCellsRange(Worksheet ws, string startCell, string endCell)
        {
            if (startCell == endCell)
                return new List<string>().Add("" + ws.Range[startCell].Value);

            return ((Array)ws.Range[startCell + ":" + endCell].Cells.Value).OfType<object>().Select(o => o.ToString()).ToList();
        }
    }
}
