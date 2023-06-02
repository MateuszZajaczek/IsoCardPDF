using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
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

            // Przepisanie wartości i formatu z komórki A1 z oryginalnego pliku do komórki B2 nowego pliku



            int targetColumn = 0;
            string lastKey = null;
            Dictionary<string, List<string>> keyValueDict = new Dictionary<string, List<string>>();

            for (int i = 0; i < 100; i++)
            {
                string orderNumber = sourceExcel.ReadCell(i, 0); // Klucz - numer zamówienia
                string profileName = sourceExcel.ReadCell(i, 1); // Nazwa profilu
                string dimensions = sourceExcel.ReadCell(i, 2); // Wymiary
                string quantity = sourceExcel.ReadCell(i, 3); // Ilość

                if (!string.IsNullOrEmpty(orderNumber))
                {
                    // Jeśli klucz (numer zamówienia) się zmienił, zwiększamy targetColumn o 4 i zapamiętujemy nowy klucz
                    if (orderNumber != lastKey)
                    {
                        targetColumn += 4;
                        lastKey = orderNumber;
                        keyValueDict[lastKey] = new List<string>();
                    }

                    // Jeśli wartość (nazwa profilu) nie była wcześniej zapisana dla danego zamówienia, przepisujemy ją i zapamiętujemy
                    if (!keyValueDict[lastKey].Contains(profileName))
                    {
                        targetExcel.WriteToCell(9 , targetColumn, profileName);
                        targetExcel.WriteToCell(10 , targetColumn, dimensions);
                        targetExcel.WriteToCell(11 , targetColumn, quantity);
                        keyValueDict[lastKey].Add(profileName);
                    }
                }
            }





            // Zapisanie i zamknięcie nowego pliku
            targetExcel.Save();
            targetExcel.Close();

            // Zamknięcie oryginalnego pliku
            sourceExcel.Close();

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
