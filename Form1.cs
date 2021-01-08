using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Cloud.Dialogflow.V2;

namespace GeneratePowerPoint
{
    public partial class Form1 : Form
    {
        private string PowerPointPath { get; set; } // file path
        private List<string> Images { get; set; } // urls to images

        public Form1()
        {
            InitializeComponent();
        }

        private void loadPPT_Click(object sender, EventArgs e)
        {
            // open file explorer dialog box
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // user selects file
                PowerPointPath = openFileDialog1.FileName;
                pptFilePath.Text = $"File: {PowerPointPath}";
            }
            else
            {
                pptFilePath.Text = $"File: n/a";
            }
        }

        // TODO: insert new slide in powerpoint
        private void generateSlide_Click(object sender, EventArgs e)
        {
            // file path has been made
            if(PowerPointPath != null)
            {
                // new slide
                // open ppt
                // add slide
                // save
                // close ppt
            }
            else
            {
                MessageBox.Show("Powerpoint file path missing", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // TODO: request new pictures
        private void titleBox_TextChanged(object sender, EventArgs e)
        {
            
        }

        // TODO: if bold text request new pictures
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void titleBox_Click(object sender, EventArgs e)
        {
            if (titleBox.Text == "Click to add title")
            {
                titleBox.Text = "";
            }
        }

        private void titleBox_Leave(object sender, EventArgs e)
        {
            if (titleBox.Text == "")
            {
                titleBox.Text = "Click to add title";
            }
        }

        private void richTextBox1_Click(object sender, EventArgs e)
        {
            if (richTextBox1.Text == "Click to add text")
            {
                richTextBox1.Text = "";
            }
        }

        private void richTextBox1_Leave(object sender, EventArgs e)
        {
            if (richTextBox1.Text == "")
            {
                richTextBox1.Text = "Click to add text";
            }
        }

        private void boldButton_Click(object sender, EventArgs e)
        {
            Font currentFont = richTextBox1.SelectionFont;
            FontStyle newFontStyle;

            if (richTextBox1.SelectionFont.Bold == true)
            {
                newFontStyle = FontStyle.Regular;
            }
            else
            {
                newFontStyle = FontStyle.Bold;
            }

            richTextBox1.SelectionFont = new Font(
               currentFont.FontFamily,
               currentFont.Size,
               newFontStyle
            );

            richTextBox1.Focus();
        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            // Close the form.
            Close();
        }

        // Example - load picture url
        //private void button1_Click(object sender, EventArgs e)
        //{
        //    pictureBox1.Load("https://i.imgur.com/AeruNiR.jpg");
        //}
    }
}
