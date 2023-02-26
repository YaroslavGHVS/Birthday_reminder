using Microsoft.Office.Interop.Excel;

namespace Birthday_reminder
{
    public class ExcelImporter
    {
        //static string excelFilePath = @"C:\Users\yokvachuk\Desktop\Files\Enternainwork\Arbeit\41. Birthday_reminder\Birthday_reminder\BirthdayFile4.xlsx";
        private string excelFilePath = @"E:\Birthday_Application\Birthday_reminder\BirthdayFile4.xlsx";
        private Application appInstance;
        private Worksheet worksheet;

        public ExcelImporter()
        {
            appInstance = new Application();
        }

        ~ExcelImporter()
        {
            appInstance.Workbooks.Close();
        }

        public ref Worksheet GetFirstWorksheet()
        {
            Workbook pivotTableWorkbook = appInstance.Workbooks.Open(excelFilePath);
            worksheet = pivotTableWorkbook.Worksheets["Birthdays"];
            return ref worksheet;
        }


    }
}
