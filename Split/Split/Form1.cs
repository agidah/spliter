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

        public Form1()
        {
            InitializeComponent();
            btnSelectFile.Click += BtnSelectFile_Click;
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

        private void ResetForm()
        {
            txtRowCount.Text = string.Empty;
            txtStartRow.Text = string.Empty;
            progressBar1.Value = 0;
            lblSelectedFile.Text = string.Empty;
        }

        private async void BtnProcess_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFilePath) || string.IsNullOrEmpty(txtRowCount.Text) || string.IsNullOrEmpty(txtStartRow.Text))
            {
                MessageBox.Show("Please select an Excel file and enter the number of rows per file and the data start row.");
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

            progressBar1.Value = 0;
            progressBar1.Maximum = (int)Math.Ceiling((double)(ReadRowCount(selectedFilePath) - startRow + 1) / rowCount);

            await Task.Run(() => SplitAndSaveExcel(selectedFilePath, startRow, rowCount));
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

                        string newFileName = Path.GetFileNameWithoutExtension(filePath) + $"_{fileIndex}.xlsx";
                        string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
                        newWorkbook.SaveAs(newFilePath);
                        fileIndex++;

                        // Update progress bar
                        this.Invoke(new Action(() =>
                        {
                            progressBar1.Value++;
                        }));
                    }
                }
            }
        }
    }
}
