using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Microsoft.Office.Interop.PowerPoint;
using System.Reflection;
using System.Drawing;

namespace JanusPPTAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {


        }
        
        private void checkedIfColorSchemeExistsElseMoveIt(string path)
        {
            if (System.IO.File.Exists(path))
            {
                //MessageBox.Show("ColorScheme Exists!.");
                if (Globals.ThisAddIn.Application.ActivePresentation != null) {
                    Globals.ThisAddIn.Application.ActivePresentation.ApplyTheme(path);
                }
                
                return;
            }
            else
            {
                //File not where we need it, so we need to copy it there!
                MessageBox.Show("Bitte gib ein, wo du das JanusTheme.potx gespeichert hast...");
                System.Windows.Forms.OpenFileDialog ofd;
                ofd = new System.Windows.Forms.OpenFileDialog();

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string sourceFile = ofd.FileName;
                    //Kopiere sourceFile to getPathToSlideBib()
                    MessageBox.Show("Versuche die Datei von " + sourceFile + " nach " + path + " zu kopieren. Sollte es dann immer noch nicht laufen, bitte von selbst machen!");
                    checkIfFolderExistsElseCreateIt(path);
                    System.IO.File.Copy(sourceFile, path);
                    if (System.IO.File.Exists(path))
                    {
                        if (Globals.ThisAddIn.Application.ActivePresentation != null)
                        {
                            Globals.ThisAddIn.Application.ActivePresentation.ApplyTheme(path);
                        }
                        MessageBox.Show("Datei von " + sourceFile + " nach " + path + " kopiert. Jetzt sollte es laufen ;)");
                    }else
                    {
                        MessageBox.Show("Das hat leider nicht geklappt :/ Kopiere bitte selber die Datei von " + sourceFile + " nach " + path + " und starte Powerpoint neu, dann sollte es laufen ;)");
                    }   
                }
                else
                {
                    //MessageBox.Show("ofd Failed");
                }
                return;
            }
            
        }

        private void checkedIfSlideBibExistsElseMoveIt(string path)
        {
            if (System.IO.File.Exists(path))
            {
                Globals.ThisAddIn.Application.ActivePresentation.ApplyTheme(path);
                return;
            }
            else
            {
                //File not where we need it, so we need to copy it there!
                MessageBox.Show("Bitte gib ein, wo du die JanusSlideBibliothek.pptx gespeichert hast...");
                System.Windows.Forms.OpenFileDialog ofd;
                ofd = new System.Windows.Forms.OpenFileDialog();

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string sourceFile = ofd.FileName;
                    //Kopiere sourceFile to getPathToSlideBib()
                    MessageBox.Show("Versuche die Datei von " + sourceFile + " nach " + path + " zu kopieren. Sollte es dann immer noch nicht laufen, bitte von selbst machen!");
                    checkIfFolderExistsElseCreateIt(path);
                    System.IO.File.Copy(sourceFile, path);
                    if (System.IO.File.Exists(path))
                    {
                        Globals.ThisAddIn.Application.ActivePresentation.ApplyTheme(path);
                        MessageBox.Show("Datei von " + sourceFile + " nach " + path + " kopiert. Jetzt sollte es laufen ;)");
                    }
                    else
                    {
                        MessageBox.Show("Das hat leider nicht geklappt :/ Kopiere bitte selber die Datei von " + sourceFile + " nach " + path + " und starte Powerpoint neu, dann sollte es laufen ;)");
                    }
                }
                else
                {
                    //MessageBox.Show("ofd Failed");
                }
                return;
            }

        }
        private void checkIfFolderExistsElseCreateIt(string pathToFile)
        {
            string pathToData = AppDomain.CurrentDomain.BaseDirectory;
            //Debug.WriteLine("pathToData:" + pathToData);

            //checked if a Resources Folder Exists
            string folderToResources = pathToData + "\\Resources";

            string fullPath = Path.GetFullPath(pathToFile).TrimEnd(Path.DirectorySeparatorChar);
            string projectName = Path.GetFileName(fullPath);
            string destinationFolder = fullPath.Replace(projectName, "");
            destinationFolder = Path.GetFullPath(destinationFolder).TrimEnd(Path.DirectorySeparatorChar);
            //MessageBox.Show("destination Folder:" + destinationFolder);
            System.IO.Directory.CreateDirectory(destinationFolder);
            Debug.WriteLine(projectName);
        }

       

        private string getPathToSlideBib()
        {
            string RunningPath = AppDomain.CurrentDomain.BaseDirectory;
            string FilePath = string.Format("{0}Resources\\FoliensammlungJanusConsulting.pptx", Path.GetFullPath(Path.Combine(RunningPath, @"..\..\")));
            //MessageBox.Show("getPathToSlideBib FilePath:" + FilePath);
            return FilePath;
        }

        private string getPathToTheme()
        {
            string RunningPath = AppDomain.CurrentDomain.BaseDirectory;
            string FilePath = string.Format("{0}Resources\\JanusTheme.potx", Path.GetFullPath(Path.Combine(RunningPath, @"..\..\")));
            //MessageBox.Show("getPathToTheme FilePath:" + FilePath);
            return FilePath;
        }

        private void SlideBibGallery_Click(object sender, RibbonControlEventArgs e)
        {
            //check if our resource is in place
            checkedIfSlideBibExistsElseMoveIt(getPathToSlideBib());
            //'ribbonGalleryObject' is the object created in Ribbon.Designer.cs
            RibbonDropDownItem item = SlideBibGallery.SelectedItem;

            string itemLabel = item.Label;
            importSlideFromSlideBibPerName(itemLabel);
        }

        private void ImageBib_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonDropDownItem item = ImageBib.SelectedItem;

            string itemLabel = item.Label;

            try
            {
                Image image = (Image)JanusPPTAddIn.Properties.Resources.ResourceManager.GetObject(itemLabel);
                Slide  slide = Globals.ThisAddIn.Application.ActivePresentation.Slides[1];
                image.Save(itemLabel + ".jpg");
                slide.Shapes.AddPicture2(itemLabel+".jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,0,0);
            }
            catch (Exception exception)
            {
                Debug.WriteLine("Tried to read Resource " + itemLabel + ", but it failed:" + exception.ToString());
            }
            
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            getFillColorFromSelectionAndMessageIT();
        }
        private int getFillColorFromSelectionAndMessageIT()
        {
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                //Shapes Are Selected
                Shape selectedShape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                FillFormat shapesFormat = selectedShape.Fill;
                ColorFormat shapesColorFormat = shapesFormat.ForeColor;
                int RGB = shapesColorFormat.RGB;
                displayRGB(RGB);
                return (RGB);
            }else
            {
                reserColor();
            }

            return -1;
        }

        private int getLineColorFromSelectionAndMessageIT()
        {
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                //Shapes Are Selected
                Shape selectedShape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                LineFormat shapesFormat = selectedShape.Line;
                ColorFormat shapesColorFormat = shapesFormat.ForeColor;
                int RGB = shapesColorFormat.RGB;
                displayRGB(RGB);
                return (RGB);
            }
            else
            {
                reserColor();
            }
            return -1;
        }
        private void displayRGB(int RGB)
        {
            Color c = Color.FromArgb(RGB);
            string R = c.R.ToString();
            string G = c.G.ToString();
            string B = c.B.ToString();
            
            RibbonEditBox editBoxR = this.editBox1;
            RibbonEditBox editBoxG = this.editBox2; 
            RibbonEditBox editBoxB = this.editBox3;

            //Dunno why, but B and R are switched up...
            editBoxR.Text = B;
            editBoxG.Text = G;
            editBoxB.Text = R;
        }

        private void reserColor()
        {
            RibbonEditBox editBoxR = this.editBox1;
            RibbonEditBox editBoxG = this.editBox2;
            RibbonEditBox editBoxB = this.editBox3;

            editBoxR.Text = "-";
            editBoxG.Text = "-";
            editBoxB.Text = "-";
        }

        //Working Flawlessly:
        private void importSlideFromSlideBibPerName(string name)
        {

            int slideNumber = -1;
            int.TryParse(name.Replace("Folie", ""), out slideNumber);
            if (slideNumber != -1)
            {
                try
                {
                    slideNumber--;
                    Globals.ThisAddIn.Application.ActivePresentation.Slides.InsertFromFile(getPathToSlideBib(), 1, slideNumber, slideNumber);
                }
                catch (Exception exception)
                {
                    Debug.WriteLine("Exceeption importing slides" + exception.ToString());
                }
            }

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            getLineColorFromSelectionAndMessageIT();
        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void padding_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            checkedIfColorSchemeExistsElseMoveIt(getPathToTheme());
        }
    }
}
