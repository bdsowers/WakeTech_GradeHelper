using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Compression;

namespace GradeHelper
{
    public partial class Form1 : Form
    {
        int mGradebookIndex = 1;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();
            if (result != DialogResult.OK)
                return;

            string path = dialog.SelectedPath;

            // Recursively find all the files in the extension list
            string[] extensions = extensionList.Text.Split(new char[] { ' ' });
            List<string> solutionFiles = new List<string>();

            foreach (string extension in extensions)
            {
                string scrubbedExtension = extension;
                if (!scrubbedExtension.StartsWith(".") && !scrubbedExtension.StartsWith("*"))
                    scrubbedExtension = "*." + scrubbedExtension;
                else if (!scrubbedExtension.StartsWith("*"))
                    scrubbedExtension = "*" + scrubbedExtension;
                Console.WriteLine(scrubbedExtension);
                string[] files = Directory.GetFiles(path, scrubbedExtension, SearchOption.AllDirectories);
                solutionFiles.AddRange(files);
            }

            solutionFiles.Sort();

            string[] exclusions = excludeList.Text.Split(new char[] { ' ' });

            // Add SLN files to the combo box
            foreach(string solutionFile in solutionFiles)
            {
                if (!IsExcluded(solutionFile, exclusions))
                {
                    listBox1.Items.Add(solutionFile);
                }
            }
			listBox1.Items.Add(Path.GetDirectoryName(path));
        }

        private bool IsExcluded(string name, string[] exclusions)
        {
            foreach(string exclusion in exclusions)
            {
                if (exclusion.Length == 0)
                    continue;

                if (name.Contains(exclusion))
                {
                    return true;
                }
            }

            return false;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listBox1.SelectedItem == null)
                return;

            string item = (string)listBox1.SelectedItem;
            System.Diagnostics.Process.Start(item);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            mGradebookIndex = 1;

            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();

            if (result != DialogResult.OK)
                return;

            string path = dialog.SelectedPath;
            if (!VerifyPathForUnpacking(path))
            {
                MessageBox.Show("The directory must have gradebook ZIP files downloaded from Blackboard");
                return;
            }

            MessageBox.Show("This may take a moment...");

			string[] zipFiles = Directory.GetFiles(path, "*.zip", SearchOption.AllDirectories);
			while (zipFiles != null && zipFiles.Length > 0)
			{
				UnpackZipFiles(zipFiles);

				zipFiles = Directory.GetFiles(path, "*.zip", SearchOption.AllDirectories);
			}

			MessageBox.Show("Done!");
        }

        bool VerifyPathForUnpacking(string path)
        {
            string[] files = Directory.GetFiles(path, "*.zip", SearchOption.TopDirectoryOnly);
            foreach(string file in files)
            {
                if (file.Contains("gradebook_"))
                    return true;
            }

            return false;
        }

		void UnpackZipFiles(string[] zipFiles)
		{
            
			foreach (string zipFile in zipFiles)
			{
				string path = Path.GetDirectoryName(zipFile);

				if (zipFile.Contains("gradebook_"))
				{
					// Root ZIP file
					string target = Path.Combine(path, "gradebook" + mGradebookIndex);
                    mGradebookIndex++;
					ZipFile.ExtractToDirectory(zipFile, target);

                    // Delete the accompanying txt files that have no value to us
                    string[] plainTextFiles = Directory.GetFiles(target, "*.txt", SearchOption.TopDirectoryOnly);
                    foreach (string plainTextFile in plainTextFiles)
                    {
                        File.Delete(plainTextFile);
                    }
                }
				else if (zipFile.Contains("_attempt_"))
				{
					// Top level student submission - parse out the student's name to create target directory
					int lastIndex = zipFile.IndexOf("_attempt_");
					int count = lastIndex - 1;
					int firstIndex = zipFile.LastIndexOf("_", lastIndex-1, count-2);

					if (lastIndex == -1 || firstIndex == -1 || lastIndex == firstIndex)
					{
						MessageBox.Show("Warning: the following file may not have unzipped properly: " + zipFile);
					}
					else
					{
						string studentName = zipFile.Substring(firstIndex + 1, lastIndex - firstIndex - 1);
                        try
                        {
                            string target = Path.Combine(path, studentName);
                            ZipFile.ExtractToDirectory(zipFile, target);
                        }
                        catch (Exception e)
                        {
                            listBox1.Items.Add("Problem with student: " + studentName);
                        }
					}
				}
				else
				{
                    // ZIP created by the student - assume it has the necessary
                    // folder inside it
                    try
                    {
                        ZipFile.ExtractToDirectory(zipFile, path);
                    }
                    catch(Exception e)
                    {
                        listBox1.Items.Add("Problem with student file: " + zipFile);
                    }
				}

				// Delete the zip file - it's no longer relevant
				File.Delete(zipFile);
			}
		}
    }
}
