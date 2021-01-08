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
using Aspose.Slides;
using System.IO;
using Aspose.Slides.Export;

namespace GeneratePowerPoint
{
    public partial class Form1 : Form
    {
        private string PowerPointPath { get; set; } // file path
        private Presentation pres { get; set; } // the open ppt
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
                pres = new Presentation(PowerPointPath);
                numSlides.Text = $"Slides #: {pres.Slides.Count}";
            }
        }

        // TODO: insert new slide in powerpoint
        private void generateSlide_Click(object sender, EventArgs e)
        {
            // file path has been made
            if(PowerPointPath != null)
            {
                // get layout type
                IMasterLayoutSlideCollection layoutSlides = pres.Masters[0].LayoutSlides;
                ILayoutSlide layoutSlide = layoutSlides.GetByType(SlideLayoutType.TitleAndObject) ?? layoutSlides.GetByType(SlideLayoutType.Title);
                if (layoutSlide == null)
                {
                    // The situation when a presentation doesn't contain some type of layouts.
                    // presentation File only contains Blank and Custom layout types.
                    // But layout slides with Custom types has different slide names,
                    // like "Title", "Title and Content", etc. And it is possible to use these
                    // names for layout slide selection.
                    // Also it is possible to use the set of placeholder shape types. For example,
                    // Title slide should have only Title pleceholder type, etc.
                    foreach (ILayoutSlide titleAndObjectLayoutSlide in layoutSlides)
                    {
                        if (titleAndObjectLayoutSlide.Name == "Title and Object")
                        {
                            layoutSlide = titleAndObjectLayoutSlide;
                            break;
                        }
                    }

                    if (layoutSlide == null)
                    {
                        foreach (ILayoutSlide titleLayoutSlide in layoutSlides)
                        {
                            if (titleLayoutSlide.Name == "Title")
                            {
                                layoutSlide = titleLayoutSlide;
                                break;
                            }
                        }

                        if (layoutSlide == null)
                        {
                            layoutSlide = layoutSlides.GetByType(SlideLayoutType.Blank);
                            if (layoutSlide == null)
                            {
                                layoutSlide = layoutSlides.Add(SlideLayoutType.TitleAndObject, "Title and Object");
                            }
                        }
                    }
                }

                // add slide
                var slide = pres.Slides.AddEmptySlide(layoutSlide);
                numSlides.Text = $"Slides #: {pres.Slides.Count}";

                // add title

                // Add an AutoShape of Rectangle type
                IAutoShape ashp = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 75, 150, 50);
                // Add TextFrame to the Rectangle
                ashp.AddTextFrame(" ");
                // Accessing the text frame
                ITextFrame txtFrame = ashp.TextFrame;
                // Create the Paragraph object for text frame
                IParagraph para = txtFrame.Paragraphs[0];
                // Create Portion object for paragraph
                IPortion portion = para.Portions[0];
                // Set Text
                portion.Text = "Aspose TextBox";

                // add body

                // add images

                // save
                string FilePath = Path.GetFileName(PowerPointPath);
                SaveFormat saveFormat = Path.GetExtension(PowerPointPath) == ".ppt" ? SaveFormat.Ppt : SaveFormat.Pptx;
                pres.Save(FilePath, saveFormat);
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
