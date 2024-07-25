using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ExcelSplitter
{
    public partial class Form1 : Form
    {
        private string selectedFilePath;
        private string selectedFolderPath;

        public Form1()
        {
            InitializeComponent();
            btnSelectFile.Click += BtnSelectFile_Click;
            btnSelectFolder.Click += BtnSelectFolder_Click;
            btnProcess.Click += BtnProcess_Click;
        }

        private void BtnSelectFile_Click(object sender, EventArgs e)
        {
            ResetForm();
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = openFileDialog.FileName;
                    lblSelectedFile.Text = Path.GetFileName(selectedFilePath);
                }
            }
        }

        private void BtnSelectFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFolderPath = folderBrowserDialog.SelectedPath;
                    lblSelectedFolder.Text = selectedFolderPath;
                }
            }
        }

        private void ResetForm()
        {
            txtRowCount.Text = string.Empty;
            txtStartRow.Text = string.Empty;
            progressBar1.Value = 0;
            progressBar1.Style = ProgressBarStyle.Continuous;
            lblSelectedFile.Text = string.Empty;
            lblSelectedFolder.Text = string.Empty;
            btnProcess.Enabled = true;
        }

        private async void BtnProcess_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFilePath) || string.IsNullOrEmpty(txtRowCount.Text) || string.IsNullOrEmpty(txtStartRow.Text) || string.IsNullOrEmpty(selectedFolderPath))
            {
                MessageBox.Show("Please select an Excel file, the number of rows per file, the data start row, and a folder to save the files.");
                return;
            }

            if (!int.TryParse(txtRowCount.Text, out int rowCount))
            {
                MessageBox.Show("Please enter a valid number for rows per file.");
                return;
            }

            if (!int.TryParse(txtStartRow.Text, out int startRow))
            {
                MessageBox.Show("Please enter a valid number for the data start row.");
                return;
            }

            // Disable the Process button and set the progress bar to Marquee style
            btnProcess.Enabled = false;
            progressBar1.Style = ProgressBarStyle.Marquee;

            await Task.Run(() => SplitAndSaveExcel(selectedFilePath, startRow, rowCount));

            // Enable the Process button and reset the progress bar style
            btnProcess.Enabled = true;
            progressBar1.Style = ProgressBarStyle.Continuous;
            progressBar1.Value = progressBar1.Maximum;

            MessageBox.Show("Files have been successfully created.");
        }

        private int ReadRowCount(string filePath)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                return worksheet.LastRowUsed().RowNumber();
            }
        }

        private void SplitAndSaveExcel(string filePath, int startRow, int rowCount)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var lastRow = worksheet.LastRowUsed().RowNumber();

                int fileIndex = 1;
                int totalParts = (int)Math.Ceiling((double)(lastRow - startRow + 1) / rowCount);

                for (int start = startRow; start <= lastRow; start += rowCount)
                {
                    using (var newWorkbook = new XLWorkbook())
                    {
                        var newWorksheet = newWorkbook.AddWorksheet("서식");

                        // Copy the header rows
                        for (int row = 1; row < startRow; row++)
                        {
                            for (int col = 1; col <= worksheet.LastColumnUsed().ColumnNumber(); col++)
                            {
                                var cell = worksheet.Cell(row, col);
                                var newCell = newWorksheet.Cell(row, col);
                                newCell.Value = cell.Value;
                                newCell.Style = cell.Style;
                                if (cell.HasFormula)
                                {
                                    newCell.FormulaA1 = cell.FormulaA1;
                                }
                            }
                            newWorksheet.Row(row).Height = worksheet.Row(row).Height;
                        }

                        // Copy merged cells in header rows
                        foreach (var merge in worksheet.MergedRanges)
                        {
                            if (merge.FirstRow().RowNumber() < startRow)
                            {
                                newWorksheet.Range(merge.RangeAddress.FirstAddress.ToString(), merge.RangeAddress.LastAddress.ToString()).Merge();
                            }
                        }

                        // Copy the data rows
                        int newRow = startRow;
                        for (int row = start; row < start + rowCount && row <= lastRow; row++, newRow++)
                        {
                            for (int col = 1; col <= worksheet.LastColumnUsed().ColumnNumber(); col++)
                            {
                                var cell = worksheet.Cell(row, col);
                                var newCell = newWorksheet.Cell(newRow, col);
                                newCell.Value = cell.Value;
                                newCell.Style = cell.Style;
                                if (cell.HasFormula)
                                {
                                    newCell.FormulaA1 = cell.FormulaA1;
                                }
                            }
                            newWorksheet.Row(newRow).Height = worksheet.Row(row).Height;
                        }

                        // Copy merged cells in data rows
                        foreach (var merge in worksheet.MergedRanges)
                        {
                            if (merge.FirstRow().RowNumber() >= start && merge.FirstRow().RowNumber() < start + rowCount)
                            {
                                int mergeStartRow = merge.FirstRow().RowNumber() - start + startRow;
                                int mergeEndRow = merge.LastRow().RowNumber() - start + startRow;
                                newWorksheet.Range(merge.FirstColumn().ColumnLetter() + mergeStartRow, merge.LastColumn().ColumnLetter() + mergeEndRow).Merge();
                            }
                        }

                        // Copy column widths
                        for (int col = 1; col <= worksheet.LastColumnUsed().ColumnNumber(); col++)
                        {
                            newWorksheet.Column(col).Width = worksheet.Column(col).Width;
                        }

                        string newFileName = Path.GetFileNameWithoutExtension(filePath) + $"_{fileIndex}.xlsx";
                        string newFilePath = Path.Combine(selectedFolderPath, newFileName);
                        newWorkbook.SaveAs(newFilePath);
                        fileIndex++;

                        // Update progress bar
                        this.Invoke(new Action(() =>
                        {
                            progressBar1.Style = ProgressBarStyle.Continuous;
                            progressBar1.Value = Math.Min(progressBar1.Maximum, (int)((double)fileIndex / totalParts * progressBar1.Maximum));
                        }));
                    }
                }
            }
        }
    }
}
