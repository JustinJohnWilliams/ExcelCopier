using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Action = System.Action;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExcelCopier
{
    public partial class Form1 : Form
    {
        // ReSharper disable InconsistentNaming
        private const string ErrorMessageTitle = "Error!";
        private const string ErrorMessageClientRequired = "Client required";
        private const string ErrorMessageNoFileSet = "No filename set";
        private const string ErrorMessageNoFileExists = "No file exists";

        private const string SuccessMessageTitle = "Success!";
        private const string SuccessMessageFilesCreated_Format = "{0} files created successfully";

        private const string ExcelFilter = "Excel Files|*.xls;*.xlsx;*.xlsm";
        // ReSharper restore InconsistentNaming

        public Form1()
        {
            InitializeComponent();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var maxCount = Convert.ToInt32(ConfigurationManager.AppSettings["MaxCount"]);
            var name = ConfigurationManager.AppSettings["Name"];

            if (string.IsNullOrEmpty(txtBoxClientName.Text))
            {
                MessageBox.Show(ErrorMessageClientRequired, ErrorMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var openFileDialog = new OpenFileDialog();
            openFileDialog.InitializeLifetimeService();
            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = ExcelFilter;
            openFileDialog.RestoreDirectory = true;

            var results = new Dictionary<string, List<string>>();
            Invoke((Action)(() => openFileDialog.ShowDialog()));

            backgroundWorker1.ReportProgress(1);

            try
            {
                var excelFile = openFileDialog.FileName;

                if (string.IsNullOrEmpty(excelFile))
                {
                    MessageBox.Show(ErrorMessageNoFileSet, ErrorMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!File.Exists(excelFile))
                {
                    MessageBox.Show(ErrorMessageNoFileExists, ErrorMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var xlApp = new Application();
                var xlWorkBook = xlApp.Workbooks.Open(excelFile);
                var xlWorkSheet = (Worksheet) xlWorkBook.Worksheets.Item[1];

                var range = xlWorkSheet.UsedRange;

                var current = 0;
                var fileCount = 0;

                do
                {
                    var amountToGrab = Math.Min(range.Rows.Count, maxCount);
                    var list = new List<string>();

                    for (var rowCount = 1; rowCount <= amountToGrab; rowCount++)
                    {
                        var resultSet = range.Cells[rowCount + current, 1] as Range;
                        if (resultSet != null)
                        {
                            list.Add((string) resultSet.Value2);
                        }
                    }

                    results.Add(
                        "{0}_{1}_{2}.txt".FormatWith(txtBoxClientName.Text, DateTime.Now.Ticks, fileCount++), list);

                    current += amountToGrab;

                    var percent = (current/(decimal) range.Rows.Count)*100;
                    backgroundWorker1.ReportProgress((int) percent);

                } while (current < range.Rows.Count);

                var resultsDirectory = "{0}\\{1}\\{2}\\{3}".FormatWith(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop), 
                    name, 
                    txtBoxClientName.Text,
                    DateTime.Now.Ticks);

                if (!Directory.Exists(resultsDirectory))
                {
                    Directory.CreateDirectory(resultsDirectory);
                }

                foreach (var file in results)
                {
                    var fileName = "{0}\\{1}".FormatWith(resultsDirectory, file.Key);
                    if (!File.Exists(fileName))
                    {
                        using (var writer = new StreamWriter(fileName))
                        {
                            foreach (var result in file.Value)
                            {
                                writer.WriteLine(result);
                            }
                        }
                    }
                }

                MessageBox.Show(SuccessMessageFilesCreated_Format.FormatWith(fileCount), SuccessMessageTitle,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                xlWorkBook.Close(false);
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ErrorMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }
    }
}