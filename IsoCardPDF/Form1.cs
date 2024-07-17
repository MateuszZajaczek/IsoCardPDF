using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace IsoCardPDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(264, 261);
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
        }



        private void Form1_Load(object sender, EventArgs e)
        {
            // Path to the program home folder. There suppose to be both: orderList and the template.
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var sourceExcel = new Excel(System.IO.Path.Combine(baseDirectory, "Rozpiska.xlsx"), 1);
            var targetExcel = new Excel(System.IO.Path.Combine(baseDirectory, "Zeszyt1.xlsx"), 1);

            int currentColumn = 2;
            int currentRow = 1;
            string lastOrderName = sourceExcel.ReadCell(2, 0);
            string lastOrderDate = null;
            string lastProfileType = null;


            _Excel._Worksheet sourceWorksheet = (_Excel._Worksheet)sourceExcel.GetWorksheet(1);
            int lastRow = sourceWorksheet.UsedRange.Rows.Count;

            // Metody ReadCell, oraz WriteCell. 1 argument - Wiersz, 2 argument - Kolumna.

            string cell = sourceExcel.ReadCell(2, 0);

            // Algorytm do walidacji danych z rozpiski.

            int countOrders = 1;

            for (int i = 0; i < lastRow; i++)
            {
                int countProductNameRows = 0;
                string orderDate = lastOrderDate;
                string orderName = sourceExcel.ReadCell(i + 1, 0);
                string productName = sourceExcel.ReadCell(i, 1);
                string profileType = lastProfileType;
                string dimensions = sourceExcel.ReadCell(i, 2);
                string quantity = sourceExcel.ReadCell(i, 3);
                string warnings = sourceExcel.ReadCell(i, 4);
                if (IsDate(orderName) || IsNumericDate(orderName))
                {
                    orderDate = orderName; // Assign the date from the order name
                    lastOrderDate = orderDate; // Update lastOrderDate with the current dateW

                    continue;
                }

                

                if (!string.IsNullOrEmpty(orderName))
                {

                    if (countOrders % 4 == 0) currentColumn--;

                    if (orderName != lastOrderName)
                    {
                        currentColumn += 5;
                        countOrders++;
                    }
                    targetExcel.WriteToCell(currentRow, currentColumn, orderName);
                    targetExcel.WriteToCell(currentRow + 1, currentColumn, orderDate);
                    lastOrderName = orderName;




                }


            }
            //    // Wypisuje wszystkie możliwe nazwy profili, ich wymiary, oraz ilości.

            //    targetExcel.WriteToCell(currentRow, targetColumn, profileName);
            //    targetExcel.WriteToCell(currentRow, targetColumn + 1, dimensions);
            //    targetExcel.WriteToCell(currentRow, targetColumn + 2, quantity);
            //    currentRow++;
            //}

            // Saving created file with orders

            string savePath = (System.IO.Path.Combine(baseDirectory, "BelkiNowe.xlsx"));
            try
            {
                targetExcel.SaveAs(savePath);
                //  Closing new file
                targetExcel.Close();

                // Closing original file
                sourceExcel.Close();
            }

            catch
            {
                Console.WriteLine("Plik nie został zapisany.");
            }
        }

        private bool IsDate(string input)
        {
            DateTime dateValue;
            return DateTime.TryParse(input, out dateValue);
        }

        private bool IsNumericDate(string input)
        {
            double number;
            if (double.TryParse(input, out number))
            {
                // Excel dates start from January 1, 1900 (serial number 1).
                // Considering dates until now, for safety we assume 1900-01-01 (1) to some future limit (e.g., 99999).
                if (number >= 1 && number <= 99999)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
