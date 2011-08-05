using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;


namespace FuzzOpen
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Executable files (*.exe)|*.exe|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            //openFileDialog1.RestoreDirectory = true;

            openFileDialog1.ShowDialog();

            listBox2.Items.Add(openFileDialog1.FileName);
            listBox2.Items.Add(Path.GetFileNameWithoutExtension(openFileDialog1.FileName));

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Process applicationtest = new Process();
            applicationtest.StartInfo.FileName = listBox2.Items[0].ToString();
            
            foreach (String arg in listBox1.Items)
            {
                applicationtest.StartInfo.Arguments = arg;
                applicationtest.Start();

                System.Threading.Thread.Sleep(Convert.ToInt32(textBoxTimer.Text) * 1000);

                /*
                System.Diagnostics.Process endProcess = new System.Diagnostics.Process();
                System.Diagnostics.ProcessStartInfo stratInfo = new System.Diagnostics.ProcessStartInfo();
                stratInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                stratInfo.FileName = "cmd.exe";
                stratInfo.Arguments = ("taskkill /F /IM " + listBox2.Items[1].ToString() + ".exe");
                endProcess.StartInfo = stratInfo;
                endProcess.Start();
                */
                
                foreach (Process clsProcess in Process.GetProcesses())
                {
                    if (clsProcess.ProcessName.StartsWith(listBox2.Items[1].ToString()))
                    {
                        System.Threading.Thread.Sleep(Convert.ToInt32(textBoxTimer.Text) * 1000);
                        applicationtest.Kill();

                        textBox3.Text += (arg + System.Environment.NewLine + "CLOSED" + System.Environment.NewLine + System.Environment.NewLine);

                        /*
                        if (applicationtest.HasExited)
                        {
                            textBox3.Text += (arg + " CLOSED" + System.Environment.NewLine);
                        }
                        else
                        {
                            //textBox3.Text += (applicationtest.StandardError.ToString() + System.Environment.NewLine);
                        }
                        */
                    }
                }
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            DialogResult result = folderBrowserDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                //Obtain information about the path
                DirectoryInfo selectedPath = new DirectoryInfo(folderBrowserDialog1.SelectedPath);

                //Check if there are files then add a label
                if (selectedPath.GetFiles().Length > 0)
                {
                    //Show all the directories using the ListBox control
                    foreach (FileInfo file in selectedPath.GetFiles())
                    {
                        listBox1.Items.Add(selectedPath + @"\" + file.Name);
                    }
                }
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBoxTimer.Text = "0";
            textBox7.Enabled = false;
            textBox8.Enabled = false;
            textBox9.Enabled = false;
        }

        private void createPDF(String tempfilename, String newPath, Random random)
        {
            //MessageBox.Show("PDF");

            string newFileName = tempfilename + ".pdf";
            newPath = System.IO.Path.Combine(newPath, newFileName);
            int sizeofwrite = random.Next(1, 10000);

            if (!System.IO.File.Exists(newPath))
            {

                // step 1: creation of a document-object
                iTextSharp.text.Document myDocument = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate());

                try
                {
                    // step 2:
                    // Now create a writer that listens to this doucment and writes the document to desired Stream.
                    iTextSharp.text.pdf.PdfWriter.GetInstance(myDocument, new FileStream(newPath, FileMode.Create));

                    // step 3:  Open the document now using
                    myDocument.Open();

                    for (int x = 0; x < sizeofwrite; x++)
                    {
                        Byte[] b = new Byte[1];
                        random.NextBytes(b);

                        // step 4: Now add some contents to the document
                        myDocument.Add(new iTextSharp.text.Paragraph(b[0].ToString()));
                    }
                }
                catch (iTextSharp.text.DocumentException de)
                {
                    Console.Error.WriteLine(de.Message);
                }
                catch (IOException ioe)
                {
                    Console.Error.WriteLine(ioe.Message);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine(ex.Message);
                }
                finally
                {
                    // step 5: Remember to close the documnet
                    myDocument.Close();
                    textBox4.AppendText("Created file: " + newPath.ToString() + System.Environment.NewLine + "SIZE: " + sizeofwrite.ToString() + "\r\n\r\n");
                }
            }
        }

        private void createMP3(String tempfilename, String newPath, Random random)
        {
            //MessageBox.Show("MP3");

            string newFileName = tempfilename + ".mp3";
            newPath = System.IO.Path.Combine(newPath, newFileName);
            int sizeofwrite = random.Next(1, 10000);

            if (!System.IO.File.Exists(newPath))
            {
            }
        }

        private void createDOC(String tempfilename, String newPath, Random random)
        {
            //MessageBox.Show("DOC");

            string newFileName = tempfilename + ".doc";
            newPath = System.IO.Path.Combine(newPath, newFileName);
            int sizeofwrite = random.Next(1, 10000);


            if (!System.IO.File.Exists(newPath))
            {
                StringBuilder sb = new StringBuilder();

                object fileName = newPath;
                object missing = System.Reflection.Missing.Value;
                object endOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

                //Start Word and create a new document.  
                Word._Application word = new Word.Application();
                Word._Document document = word.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                for (int x = 0; x < sizeofwrite; x++)
                {
                    Byte[] b = new Byte[1];
                    random.NextBytes(b);

                    sb.AppendLine(b[0].ToString());
                }

                word.Selection.TypeText(sb.ToString());

                document.SaveAs2(ref fileName, ref missing, ref missing, ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing, ref missing);

                document.Close(ref missing, ref missing, ref missing);
                word.Quit(ref missing, ref missing, ref missing);

                textBox4.AppendText("Created file: " + newPath.ToString() + System.Environment.NewLine + "SIZE: " + sizeofwrite.ToString() + "\r\n\r\n");
            }
        }

        private void createDOCX(String tempfilename, String newPath, Random random)
        {
            //MessageBox.Show("DOCX");
           
            string newFileName = tempfilename + ".docx";
            newPath = System.IO.Path.Combine(newPath, newFileName);
            int sizeofwrite = random.Next(1, 10000);


            if (!System.IO.File.Exists(newPath))
            {
                StringBuilder sb = new StringBuilder();
           
                object fileName = newPath;
                object missing = System.Reflection.Missing.Value;
                object endOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

                //Start Word and create a new document.  
                Word._Application word = new Word.Application();
                Word._Document document = word.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                for (int x = 0; x < sizeofwrite; x++)
                {
                    Byte[] b = new Byte[1];
                    random.NextBytes(b);

                    sb.AppendLine(b[0].ToString());
                }

                word.Selection.TypeText(sb.ToString());

                /*
                //Insert a paragraph at the beginning of the document.
                Word.Paragraph para1;
                para1 = document.Content.Paragraphs.Add(ref missing);
                para1.Range.Text = "Adding paragraph thru C#";
                para1.Range.Font.Bold = 1;
                para1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
                para1.Range.InsertParagraphAfter();
                */
                //word.Visible = true;

                document.SaveAs2(ref fileName, ref missing, ref missing, ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, 
                                    ref missing, ref missing, ref missing, ref missing);

                document.Close(ref missing, ref missing, ref missing);
                word.Quit(ref missing, ref missing, ref missing);

                textBox4.AppendText("Created file: " + newPath.ToString() + System.Environment.NewLine + "SIZE: " + sizeofwrite.ToString() + "\r\n\r\n");
            }
        }

        private void createXLS(String tempfilename, String newPath, Random random)
        {
            //MessageBox.Show("XLS");

            string newFileName = tempfilename + ".xls";
            newPath = System.IO.Path.Combine(newPath, newFileName);
            int sizeofwrite = random.Next(1, 10000);

            if (!System.IO.File.Exists(newPath))
            {
            }
        }

        private void createXLSX(String tempfilename, String newPath, Random random)
        {
            //MessageBox.Show("XLSX");

            string newFileName = tempfilename + ".xlsx";
            newPath = System.IO.Path.Combine(newPath, newFileName);
            int sizeofwrite = random.Next(1, 10000);

            if (!System.IO.File.Exists(newPath))
            {
            }
        }

        private void createPPT(String tempfilename, String newPath, Random random)
        {
            //MessageBox.Show("PPT");

            string newFileName = tempfilename + ".ppt";
            newPath = System.IO.Path.Combine(newPath, newFileName);
            int sizeofwrite = random.Next(1, 10000);

            if (!System.IO.File.Exists(newPath))
            {
            }
        }

        private void createPPTX(String tempfilename, String newPath, Random random)
        {
            //MessageBox.Show("PPTX");

            string newFileName = tempfilename + ".pptx";
            newPath = System.IO.Path.Combine(newPath, newFileName);
            int sizeofwrite = random.Next(1, 10000);

            if (!System.IO.File.Exists(newPath))
            {
            }
        }

        private void createAVI(String tempfilename, String newPath, Random random)
        {
            //MessageBox.Show("AVI");

            string newFileName = tempfilename + ".avi";
            newPath = System.IO.Path.Combine(newPath, newFileName);
            int sizeofwrite = random.Next(1, 10000);

            if (!System.IO.File.Exists(newPath))
            {
            }
        }

        private String creatFileName(String newPath, Random random)
        {
            System.IO.Directory.CreateDirectory(newPath);
            int filenamesize = random.Next(100);
            String tempfilename = "";
            int[] forbidden = { 47, 92, 58, 42, 63, 34, 60, 62, 124 }; //list of forbidden characters to use on filename

            for (int x = 0; x < filenamesize; x++)
            {
                int randomchar = random.Next(32, 126);

                if (forbidden.Contains(randomchar))
                {
                    x--;
                }
                else
                {
                    char nextchar = Convert.ToChar(randomchar);
                    tempfilename = tempfilename + nextchar.ToString();
                }
            }

            return tempfilename;
        }


        private void button4_Click(object sender, EventArgs e)
        {            
            if (textBox5.Text == "")
            {
                textBox4.Text = "PROJECT NAME MISSING";
            }
            else
            {
                int numberoffiles = Convert.ToInt32(textBox2.Text);
                
                if(numberoffiles > 10){
                    System.Windows.Forms.MessageBox.Show("Warning! Depending on your CPU this might take a while!");
                }

                Random random = new Random();
                string activeDir = textBox1.Text.ToString();
                string projectname = textBox5.Text.ToString();
                string newPath = System.IO.Path.Combine(activeDir, projectname);

                foreach (int indexChecked in checkedListBox1.CheckedIndices)
                {
                    if (checkedListBox1.GetItemCheckState(indexChecked) == CheckState.Checked)
                    {
                        String tempfilename;

                        for (int i = 0; i < numberoffiles; i++)
                        {
                            tempfilename = creatFileName(newPath, random);

                            if (indexChecked == 0)
                            {
                                createPDF(tempfilename, newPath, random);
                            }
                            else if (indexChecked == 1)
                            {
                                createMP3(tempfilename, newPath, random);
                            }
                            else if (indexChecked == 2)
                            {
                                createDOC(tempfilename, newPath, random);
                            }
                            else if (indexChecked == 3)
                            {
                                createDOCX(tempfilename, newPath, random);
                            }
                            else if (indexChecked == 4)
                            {
                                createXLS(tempfilename, newPath, random);
                            }
                            else if (indexChecked == 5)
                            {
                                createXLSX(tempfilename, newPath, random);
                            }
                            else if (indexChecked == 6)
                            {
                                createPPT(tempfilename, newPath, random);
                            }
                            else if (indexChecked == 7)
                            {
                                createPPTX(tempfilename, newPath, random);
                            }
                            else if (indexChecked == 8)
                            {
                                createAVI(tempfilename, newPath, random);
                            }
                        }
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            DialogResult result = folderBrowserDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                //Obtain information about the path
                DirectoryInfo selectedpathforfiles = new DirectoryInfo(folderBrowserDialog1.SelectedPath);
                textBox1.Text = selectedpathforfiles.ToString();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= 126; i++)
            {
                char nextchar = Convert.ToChar(i);
                textBox4.AppendText("NUMBER: " + i + "CHAR: " + nextchar.ToString());
                textBox4.AppendText("\n");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            int numberofcharacters = Convert.ToInt32(textBox11.Text);
            String filename = textBox10.Text.ToString() + textBox12.Text.ToString();
            string activeDir = textBox14.Text.ToString();
            String newPath = System.IO.Path.Combine(activeDir, filename);
            textBox13.AppendText(newPath);
            char chartowrite = Convert.ToChar(textBox15.Text.ToString());
            byte towrite = Convert.ToByte(chartowrite);
            if (!System.IO.File.Exists(newPath))
            {
                using (System.IO.FileStream fs = System.IO.File.Create(newPath))
                {
                    textBox13.AppendText("FILE BEING CREATED: " + newPath);
                    for (int i = 0; i < numberofcharacters; i++)
                    {
                        fs.WriteByte(towrite);
                    }
                }
            }

        }

        private void button9_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            DialogResult result = folderBrowserDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                //Obtain information about the path
                DirectoryInfo selectedpathforfiles = new DirectoryInfo(folderBrowserDialog1.SelectedPath);
                textBox14.AppendText(selectedpathforfiles.ToString());
            }
        }
    }
}


