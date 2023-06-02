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

            int targetColumn = 0;
            string lastOrderNumber = null;
            Dictionary<string, List<string>> orderProfileDict = new Dictionary<string, List<string>>();
            Dictionary<string, int> profileDimensionsDict = new Dictionary<string, int>();

            for (int i = 1; i < 100; i++)
            {
                string orderNumber = sourceExcel.ReadCell(i, 0); // Numer zamówienia
                string profileName = sourceExcel.ReadCell(i, 1); // Nazwa profilu
                string dimensions = sourceExcel.ReadCell(i, 2); // Wymiary
                string quantity = sourceExcel.ReadCell(i, 3); // Ilość

                if (!string.IsNullOrEmpty(orderNumber))
                {
                    // Jeśli numer zamówienia się zmienił, zwiększamy targetColumn o 4 i resetujemy słowniki
                    if (orderNumber != lastOrderNumber)
                    {
                        targetColumn += 4;
                        lastOrderNumber = orderNumber;
                        orderProfileDict.Clear();
                        profileDimensionsDict.Clear();
                    }

                    string key = $"{orderNumber}-{profileName}";
                    if (!orderProfileDict.ContainsKey(key))
                    {
                        // Jeśli to jest nowy profil w danym zamówieniu, zapisujemy numer zamówienia w odpowiedniej komórce
                        targetExcel.WriteToCell(1, targetColumn + 2, orderNumber);
                        orderProfileDict[key] = new List<string>();
                    }

                    // Sprawdzamy, czy dla danego profilu w zamówieniu są dostępne dodatkowe wymiary
                    if (!profileDimensionsDict.ContainsKey(key))
                    {
                        profileDimensionsDict[key] = 0;
                    }

                    int profileRow = profileDimensionsDict[key] * 4 + 9;

                    targetExcel.WriteToCell(profileRow, targetColumn, profileName);
                    targetExcel.WriteToCell(profileRow, targetColumn + 1, dimensions);
                    targetExcel.WriteToCell(profileRow, targetColumn + 2, quantity);

                    orderProfileDict[key].Add(dimensions);
                    profileDimensionsDict[key]++;
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
