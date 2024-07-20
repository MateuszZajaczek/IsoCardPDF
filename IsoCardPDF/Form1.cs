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
using IsoCardPDF.Entities;


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
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            var sourceExcel = new Excel(System.IO.Path.Combine(baseDirectory, "Rozpiska.xlsx"), 1);
            var targetExcel = new Excel(System.IO.Path.Combine(baseDirectory, "Zeszyt.xlsx"), 1);
            string lastOrderDate = null;
            string currentOrderName = null;

            // Iso card template already created in Excel file. Cells A1:D26
            Range isoCardTemplate = targetExcel.GetWorksheet(1).Range["A1:D26"];

            _Excel._Worksheet sourceWorksheet = (_Excel._Worksheet)sourceExcel.GetWorksheet(1);
            int lastRow = sourceWorksheet.UsedRange.Rows.Count;  // Get the number of rows currently in use in the source worksheet.

            // Reading Excel data from breakdown.
            List<Order> orders = new List<Order>();
            Order currentOrder = null;

            for (int i = 0; i < lastRow; i++)
            {
                string orderName = sourceExcel.ReadCell(i, 0);
                string profileTypeStr = sourceExcel.ReadCell(i, 1);
                string dimensions = sourceExcel.ReadCell(i, 2);
                string quantityStr = sourceExcel.ReadCell(i, 3);
                string warning = sourceExcel.ReadCell(i, 4);

                // Ensure quantityStr is a valid integer
                int quantity = 0;
                if (!string.IsNullOrWhiteSpace(quantityStr))
                {
                    if (!int.TryParse(quantityStr, out quantity))
                    {
                        Console.WriteLine($"Invalid quantity format at row {i + 1}: {quantityStr}");
                        continue; // Skip this row if quantity is invalid
                    }
                }

                if (IsDate(orderName) || IsNumericDate(orderName))
                {
                    lastOrderDate = orderName;
                    continue;
                }

                if (!string.IsNullOrEmpty(orderName) && orderName != currentOrderName)
                {
                    currentOrderName = orderName;
                    currentOrder = new Order(currentOrderName); // Use a unique identifier for the order
                    orders.Add(currentOrder);
                }

                if (currentOrder == null)
                {
                    Console.WriteLine($"No current order to add position at row {i + 1}");
                    continue; // Skip this row if there is no current order
                }

                ProfileType profileType;
                if (!EnumHelper.TryParseDescription(profileTypeStr, out profileType))
                {
                    Console.WriteLine($"Invalid profile type at row {i + 1}: {profileTypeStr}");
                    continue; // Skip this row if profile type is invalid
                }

                Position position = string.IsNullOrEmpty(warning)
                    ? new Position(profileType, dimensions, quantity)
                    : new Position(profileType, dimensions, quantity, warning);

                currentOrder.AddPosition(position);
            }

            // Display order details
            foreach (var order in orders)
            {
                Console.WriteLine($"Order ID: {order.OrderId}");
                foreach (var pos in order.Positions)
                {
                    Console.WriteLine($"Profile Type: {pos.ProfileType}, Dimension: {pos.Dimension}, Quantity: {pos.Quantity}, Warning: {pos.Warning}");
                }
            }

            // Save new file
            string savePath = System.IO.Path.Combine(baseDirectory, "BelkiNowe.xlsx");
            try
            {
                targetExcel.SaveAs(savePath);
                targetExcel.Close();
                sourceExcel.Close();
            }
            catch
            {
                Console.WriteLine("Plik nie został zapisany.");
            }
        }

        public void OpenFile()
        {
            Excel excel = new Excel(@"C:\Users\Mateusz\Desktop\Github\Aplikacja do pracy\AppIso\IsoCardPDF\IsoCardPDF\Książka.xlsx", 1);
            MessageBox.Show(excel.ReadCell(0, 0));
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
