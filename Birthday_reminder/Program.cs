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
            Application app = new Application();
            app.Visible = true;
            Workbook sampleWorkbook = app.Workbooks.Add();
            //Workbook existingWorkbook = app.Workbooks.Open(@"C:\Users\yokvachuk\Desktop\Files\Enternainwork\Arbeit\41. Birthday_reminder\Birthday_reminder\BirthdayFile.xlsx");

            // declare worksheet object
            Worksheet worksheet = sampleWorkbook.Worksheets["Birthdays"];

            //change the value of one cell
            worksheet.Range["A1"].Value = "First name";
        }
    }
}

// https://www.youtube.com/watch?v=_Hn4hbe1NxM