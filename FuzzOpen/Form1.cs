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
using iTextSharp.text;
using iTextSharp.text.pdf;


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
            applicationtest.StartInfo.FileName = listBox2.Items[0].ToString();//textBox1.Text.ToString();

            foreach (String arg in listBox1.Items)
            {
                applicationtest.StartInfo.Arguments = arg;

                applicationtest.Start();

                //applicationtest.WaitForExit();

                foreach (Process clsProcess in Process.GetProcesses())
                {


                    if (clsProcess.ProcessName.StartsWith(listBox2.Items[1].ToString()))
                    {
                        //textBox3.Text += clsProcess.ToString() + System.Environment.NewLine;
                        System.Threading.Thread.Sleep(Convert.ToInt32(textBoxTimer.Text) * 1000);
                        applicationtest.Kill();
                        //clsProcess.Kill();
                        textBox3.Text += (arg + " CLOSED" + System.Environment.NewLine);
                        /*
                        if (applicationtest.HasExited == true)
                        {
                            textBox3.Text += (arg + " CLOSED" + System.Environment.NewLine);
                        }
                        else
                        {
                            //textBox3.Text += (applicationtest.StandardError.ToString() + System.Environment.NewLine);
                        }
                         * */
                    }
                }
                //System.Threading.Thread.Sleep(5000);
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            /*openFileDialog2.ShowDialog();
            textBox2.Text = openFileDialog2.FileName.ToString();*/

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
                Random random = new Random();

                for (int i = 0; i < numberoffiles; i++)
                {
                    string activeDir = textBox1.Text.ToString();
                    string projectname = textBox5.Text.ToString();
                    string newPath = System.IO.Path.Combine(activeDir, projectname);

                    System.IO.Directory.CreateDirectory(newPath);
                    int filenamesize = random.Next(100);
                    String tempfilename = "";
                    int[] forbidden = { 47, 92, 58, 42, 63, 34, 60, 62, 124 };

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

                    string newFileName = tempfilename + ".pdf";
                    newPath = System.IO.Path.Combine(newPath, newFileName);
                    int sizeofwrite = random.Next(1, 10000);

                    if (!System.IO.File.Exists(newPath))
                    {

                        // step 1: creation of a document-object
                        Document myDocument = new Document(PageSize.A4.Rotate());

                        try
                        {
                            // step 2:
                            // Now create a writer that listens to this doucment and writes the document to desired Stream.
                            PdfWriter.GetInstance(myDocument, new FileStream(newPath, FileMode.Create));

                            // step 3:  Open the document now using
                            myDocument.Open();
  
                            for (int x = 0; x < sizeofwrite; x++)
                            {
                                Byte[] b = new Byte[1];
                                random.NextBytes(b);

                                // step 4: Now add some contents to the document
                                myDocument.Add(new Paragraph(b[0].ToString()));
                                //myDocument.Add(new Paragraph("First Pdf File made by Salman using iText"));
                            }
                        }
                        catch (DocumentException de)
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
    }
}


