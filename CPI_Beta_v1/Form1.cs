using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using CPI_Beta_v1.Properties;

namespace CPI_Beta_v1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == null) return;

            var lines = File.ReadAllLines(textBox1.Text, Encoding.Default);
            var txtHandler = new TxtHandler();
            var list = txtHandler.BuildInterventions(lines);
            var excelBuilder = new ExcelBuilder();
            excelBuilder.GenerateExcel(list,dateTimePicker1.Value.Year);
        }

        private void textBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // Set filter options and filter index.
            var openFileDialog1 = new OpenFileDialog
            {
                Filter = Resources.Form1_textBox1_MouseDoubleClick_Text_Files___txt____txt,
                FilterIndex = 1
            };

            

            // Process input if the user clicked OK.
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;

            }

        }
    }
}
