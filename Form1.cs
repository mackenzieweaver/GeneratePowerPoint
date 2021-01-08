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
        public Form1()
        {
            InitializeComponent();
        }
        
        // TODO: insert new slide in powerpoint
        private void generateSlide_Click(object sender, EventArgs e)
        {
            // open file explorer dialog box
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // user selects file
                var pp = openFileDialog1.FileName;
            }
            // open and insert slide
        }

        // TODO: request new pictures
        private void titleBox_TextChanged(object sender, EventArgs e)
        {
            var hook = new WebhookResponse();
            
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

        private void closeButton_Click(object sender, EventArgs e)
        {
            // Close the form.
            Close();
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
    }
}
