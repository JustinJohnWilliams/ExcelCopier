using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCopier
{
    public partial class Form1 : Form
    {
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
                MessageBox.Show("Client required", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var openFileDialog = new OpenFileDialog();
            openFileDialog.InitializeLifetimeService();
            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.RestoreDirectory = true;

            var results = new Dictionary<string, List<string>>();
            Invoke((Action)(() => { openFileDialog.ShowDialog(); }));

            try
            {
                var excelFile = openFileDialog.FileName;

                if (string.IsNullOrEmpty(excelFile))
                {
                    MessageBox.Show("No filename set", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                if (!File.Exists(excelFile))
                {
                    MessageBox.Show("No file exists", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                var xlApp = new Excel.Application();
                var xlWorkBook = xlApp.Workbooks.Open(excelFile);
                var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                var range = xlWorkSheet.UsedRange;

                var rowCount = 0;
                var current = 0;
                var fileCount = 0;

                do
                {
                    var amountToGrab = Math.Min(range.Rows.Count, maxCount);
                    var list = new List<string>();

                    for (rowCount = 1; rowCount <= amountToGrab; rowCount++)
                    {
                        list.Add((string)(range.Cells[rowCount + current, 1] as Excel.Range).Value2);
                    }

                    results.Add(String.Format("{0}_{1}_{2}.txt", txtBoxClientName.Text, DateTime.Now.Ticks, fileCount++), list);

                    current += amountToGrab;

                    var percent = ((decimal)current / (decimal)range.Rows.Count) * 100;
                    backgroundWorker1.ReportProgress((int)percent);

                } while (current < range.Rows.Count);

                string resultsDirectory = string.Format("{0}\\{1}\\{2}\\{3}", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), name, txtBoxClientName.Text, DateTime.Now.Ticks);

                if (!Directory.Exists(resultsDirectory))
                {
                    Directory.CreateDirectory(resultsDirectory);
                }

                foreach (var file in results)
                {
                    var fileName = string.Format("{0}\\{1}", resultsDirectory, file.Key);
                    if (!File.Exists(fileName))
                    {
                        using (StreamWriter writer = new StreamWriter(fileName))
                        {
                            foreach (var result in file.Value)
                            {
                                writer.WriteLine(result);
                            }
                        }
                    }
                }

                MessageBox.Show(String.Format("{0} files created successfully", fileCount), "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                xlWorkBook.Close(false);
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
