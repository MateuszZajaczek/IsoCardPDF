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
            var sourceExcel = new Excel("Książka2.xlsx", 1);
            var targetExcel = new Excel("Zeszyt.xlsx", 1);

            int targetColumn = -4;
            string lastOrderNumber = null;
            int currentRow = 9;

            // Zakładając, że pierwsza karta ISO jest już stworzona w A1:D9
            Range isoCardTemplate = targetExcel.GetWorksheet(1).Range["A1:D9"];

            _Excel._Worksheet sourceWorksheet = (_Excel._Worksheet)sourceExcel.GetWorksheet(1);
            int lastRow = sourceWorksheet.UsedRange.Rows.Count;

            for (int i = 1; i <= lastRow; i++)
            {
                string orderNumber = sourceExcel.ReadCell(i, 0); // Numer zamówienia
                string profileName = sourceExcel.ReadCell(i, 1); // Nazwa profilu
                string dimensions = sourceExcel.ReadCell(i, 2); // Wymiary
                string quantity = sourceExcel.ReadCell(i, 3); // Ilość

                if (!string.IsNullOrEmpty(orderNumber))
                {
                    // Jeśli numer zamówienia się zmienił, zwiększamy targetColumn o 4 i resetujemy wiersz
                    if (orderNumber != lastOrderNumber)
                    {
                        targetColumn += 4;
                        lastOrderNumber = orderNumber;
                        currentRow = 8;

                        // Skopiuj kartę ISO na nową pozycję
                        isoCardTemplate.Copy(targetExcel.GetWorksheet(1).Range[targetExcel.GetColumnName(targetColumn + 1) + "1:" + targetExcel.GetColumnName(targetColumn + 4) + "9"]);

                        // Scalamy zdefiniowane zakresy komórek
                        MergeCells(targetColumn + 1, 1, targetColumn + 4, 1);
                        MergeCells(targetColumn + 1, 2, targetColumn + 2, 2);
                        MergeCells(targetColumn + 3, 2, targetColumn + 4, 2);
                        MergeCells(targetColumn + 1, 3, targetColumn + 2, 3);
                        MergeCells(targetColumn + 3, 3, targetColumn + 4, 3);
                        MergeCells(targetColumn + 1, 4, targetColumn + 2, 4);
                        MergeCells(targetColumn + 3, 4, targetColumn + 4, 4);
                        MergeCells(targetColumn + 1, 5, targetColumn + 4, 5);
                        MergeCells(targetColumn + 1, 6, targetColumn + 4, 7);
                    }



                    targetExcel.WriteToCell(1, targetColumn + 2, orderNumber);  // Wypisuje numer zamówienia do odpowiedniej kolumny.


                }
                // Wypisuje wszystkie możliwe nazwy profili, ich wymiary, oraz ilości.

                targetExcel.WriteToCell(currentRow, targetColumn, profileName);
                targetExcel.WriteToCell(currentRow, targetColumn + 1, dimensions);
                targetExcel.WriteToCell(currentRow, targetColumn + 2, quantity);
                currentRow++;
            }
            
            // Zapisanie nowego pliku
            string savePath = @"BelkiNowe.xlsx";

            try
            {
                targetExcel.SaveAs(savePath);
                // Zamknięcie nowego pliku
                targetExcel.Close();

                // Zamknięcie oryginalnego pliku
                sourceExcel.Close();
            }
            
            catch
            {
                Console.WriteLine("Plik nie został zapisany.");
            }

            

            void MergeCells(int startColumn, int startRow, int endColumn, int endRow)
            {
                var rangeToMerge = targetExcel.GetWorksheet(1).Range[targetExcel.GetColumnName(startColumn) + startRow.ToString() + ":" + targetExcel.GetColumnName(endColumn) + endRow.ToString()];
                rangeToMerge.Merge();
            }
        }



        public void OpenFile()
        {
            Excel excel = new Excel(@"C:\Users\Mateusz\Desktop\Github\Aplikacja do pracy\AppIso\IsoCardPDF\IsoCardPDF\Książka.xlsx", 1);
            MessageBox.Show(excel.ReadCell(0, 0));
        }



        public void WriteData()
        {
            Excel excel = new Excel(@"Test.xlsx", 1);
            excel.WriteToCell(0, 0, "Test2");
            excel.Save();
            excel.SaveAs(@"Test2.xlsx");
        }
    }
}
