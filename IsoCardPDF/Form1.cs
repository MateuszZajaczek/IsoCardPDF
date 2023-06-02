using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            WriteData();

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
