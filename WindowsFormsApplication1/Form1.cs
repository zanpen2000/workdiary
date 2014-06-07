using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NetOffice;
using NetOffice.ExcelApi;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = @"E:\迅雷下载\RG03_042_20140604.xls";
        }

        private void textBox1_DoubleClick(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = dialog.FileName;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            NetOffice.ExcelApi.Application excelApp = new NetOffice.ExcelApi.Application();
            excelApp.DisplayAlerts = false;
            
            NetOffice.ExcelApi.Workbook workBook = excelApp.Workbooks.Open(textBox1.Text);
            Worksheet workSheet = (Worksheet)workBook.Worksheets[1];

            Person person = new Person();
            
            person.Name = workSheet.Range("B3").Value2.ToString();
            person.Id = workSheet.Range("E3").Value2.ToString();
            person.Date = DateTime.FromOADate((double)workSheet.Range("G3").Value2).ToString();
            person.Company = workSheet.Range("C4").Value2.ToString();
            person.Department = workSheet.Range("C5").Value2.ToString();

            //MessageBox.Show(person.Name);

            excelApp.Quit();
            excelApp.Dispose(); 
        }
    }
}
