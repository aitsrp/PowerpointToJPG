using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Powerpoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Diagnostics;

namespace PowerpointToJPG
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            DragDropEffects effect = DragDropEffects.None;
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var path = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
                FileInfo inf = new FileInfo(path);
                string ext = inf.Extension.ToLower();
                if (ext == ".ppt" || ext == ".pptx")
                    effect = DragDropEffects.Copy;
            }

            e.Effect = effect;
        }

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                label1.Text = "Processing presentations";
                var filelist = ((string[])e.Data.GetData(DataFormats.FileDrop));
                textBox1.Text = "";
                progressBar1.Value = 0;
                for (var i = 0; i < filelist.Length; i++)
                {
                    string path = filelist[i];
                    FileInfo inf = new FileInfo(path);
                    progressBar1.Minimum = 0;
                    Console.WriteLine("min: " + progressBar1.Minimum);
                    progressBar1.Maximum = filelist.Length;
                    Console.WriteLine("max: " + progressBar1.Maximum);
                    Console.WriteLine("pval: " + progressBar1.Value);
                    progressBar1.PerformStep();
                    Console.WriteLine("fval: " + progressBar1.Value);
                    PowerpointToJPG(path);
                    label1.Text = "Converted " + (i+1) + " of " + (filelist.Length) + " presentations";
                    label1.Refresh();
                    if (i == filelist.Length - 1)
                    {
                        log.Text += "Process done\r\n";
                        processKill();
                    }
                }
                if (textBox1.Text != "")
                    Clipboard.SetText(textBox1.Text);
            }
        }

        private void PowerpointToJPG(string file)
        {
            FileInfo inf = new FileInfo(file);
            string folder = GetFolderName(inf);
            string directory = inf.DirectoryName + "\\" + folder;
            Console.WriteLine("Final directory: " + directory);

            if (Directory.Exists(directory))
            {
                if (MessageBox.Show("A folder with the name '" + folder + "' already exists in the current directory.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) == DialogResult.OK)
                {
                    return;
                }
            }
            else
            {
                textBox1.Text += "Processing : " + inf.Name + "\r\n";
                Powerpoint.Application pptApp = new Powerpoint.Application();
                Powerpoint.Presentation pptPresentation = pptApp.Presentations.Open(file, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                pptPresentation.Export(directory, "JPG");
            }
        }

        private string GetFolderName(FileInfo inf)
        {
            string rawname = Path.GetFileNameWithoutExtension(inf.Name);
            string finalname = "";
            if (rawname.Contains('-')) {
                finalname = rawname.Split('-')[1].Split('_')[0].Trim();
            } else if (rawname.Contains('_'))
            {
                finalname = rawname.Split('_')[1].Split('_')[0].Trim();
            }
                Console.WriteLine("finalname: " + finalname);
            return finalname;
        }
        private void processKill()
        {
            Process[] pptprocs = Process.GetProcessesByName("POWERPNT");
            foreach (Process proc in pptprocs)
            {
                proc.Kill();
            }
        }
    }
}
