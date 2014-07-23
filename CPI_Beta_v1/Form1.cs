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
            panel2.DragDrop += panel2_DragDrop;
            panel2.AllowDrop = true;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == null) return;
            try
            {
                var lines = File.ReadAllLines(textBox1.Text, Encoding.Default);
                var txtHandler = new TxtHandler();
                var list = txtHandler.BuildInterventions(lines);
                var excelBuilder = new ExcelBuilder();
                excelBuilder.GenerateExcel(list, dateTimePicker1.Value.Year);
            }
            catch (ArgumentException)
            {
                MessageBox.Show(Resources.Form1_button2_Click_O_sistema_não_consegue_interpretar_o_ficheiro_,
                    Resources.Form1_button2_Click_AVISO, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
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
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            textBox1.Text = openFileDialog1.FileName;
            panel1.BringToFront();
            pictureBox2.Image = Resources.imageTXT;
            label4.Text = openFileDialog1.FileName.Split('\\').Last();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            panel2.BringToFront();
            textBox1.Text = string.Empty;
        }

    
       

    }
}
