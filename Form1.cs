using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GeneratePowerPoint
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        // 1 TODO: export a file in .pptx
        private void generateSlide_Click(object sender, EventArgs e)
        {
            
        }

        // 2 TODO: request new pictures
        private void titleBox_TextChanged(object sender, EventArgs e)
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

        // 4 TODO: if bold text request new pictures
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

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
    }
}
