using System;
using System.IO;
using System.Linq;
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
            panel2.DragEnter += panel2_DragEnter;
            panel2.DragOver += panel2_DragOver;
            panel2.DragLeave += panel2_DragLeave;
            panel2.DragDrop += panel2_DragDrop;
            panel2.AllowDrop = true;

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

        private void Form1_Load(object sender, EventArgs e)
        {
            label3.Text = Resources.Form1_Form1_Load_;
        }



      private  void panel2_DragDrop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (files != null && File.Exists(files[0]))
            {
                var extension = Path.GetExtension(files[0]);
                if (extension != null && !extension.Equals(".txt", StringComparison.InvariantCultureIgnoreCase))
                {
                    MessageBox.Show(Resources.Form1_panel2_DragDrop_File_format_not_allowed);
                }
                else
                {
                    
                    panel1.BringToFront();
                    pictureBox2.Image = Resources.imageTXT;
                    label4.Text = files[0].Split('\\').Last();
                    textBox1.Text = files[0];
              
                }
            }
            else
            {
                MessageBox.Show(Resources.Form1_panel2_DragDrop_File_does_not_exists__);
            }
        }

      private void panel2_DragLeave(object sender, EventArgs e)
        {
         
        }

      private void panel2_DragOver(object sender, DragEventArgs e)
        {
           
          
        }

      private void panel2_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void button1_MouseClick(object sender, MouseEventArgs e)
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
