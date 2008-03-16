using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Configuration;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using ExcelGrepper;

namespace ExcelGrepper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(this.textBox1.Text))
            {
                this.textBox1.Enabled =
                this.button1.Enabled = false;
                backgroundWorker1.RunWorkerAsync();
            }
        }

        ExcelGrepDS.ExcelGrepResultDataTable table;
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            table = GrepExcels(this.textBox1.Text);
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            dataGridView1.DataSource = table;
            this.textBox1.Enabled =
            this.button1.Enabled = true;
            this.progressBar1.Value = 0;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // 進捗状況をプログレスバーに設定
            progressBar1.Value = e.ProgressPercentage;
        }

        ExcelGrepDS.ExcelGrepResultDataTable GrepExcels(string sword)
        {
            ExcelGrepDS.ExcelGrepResultDataTable table = new ExcelGrepDS.ExcelGrepResultDataTable();

            string baseDir = ConfigurationManager.AppSettings["basedir"];
            FileInfo[] fileinfos = new DirectoryInfo(baseDir)
                .GetFiles("*.xls", SearchOption.AllDirectories);
            int len = fileinfos.Length;
            int count = 0;
            foreach (FileInfo fileInfo in fileinfos)
            {
                try
                {
                    //Debug.WriteLine("Read " + fileInfo.Name);
                    ReadExcel(fileInfo, table, sword);
                    backgroundWorker1.ReportProgress(++count * 100 / len);
                }
                catch (COMException come)
                {
                    Debug.WriteLine("[error] " + come.Message);
                }
            }

            return table;
        }

        Excel.Application objApp;

        void ReadExcel(FileInfo excelFile, ExcelGrepDS.ExcelGrepResultDataTable table, string sword)
        {
            string excludeSheetNames = ConfigurationManager.AppSettings["excludeSheetNames"];

            Excel.Workbooks objBooks;

            //Open App
            objApp = new Excel.Application();
            objApp.DisplayAlerts = false;
            objBooks = objApp.Workbooks;
            //Open Book
            Excel.Workbook objBook = objBooks.Open(excelFile.FullName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //Main
            foreach (Excel.Worksheet objSheet in objBook.Worksheets)
            {
                string sheetName = objSheet.Name;
                if (excludeSheetNames.Contains(sheetName)) continue;

                Debug.WriteLine(excelFile.Name + ":" + sheetName);

                Range usedrange = objSheet.UsedRange;
                int Xmax = usedrange.Rows.Count + 1;
                int Ymax = usedrange.Columns.Count + 1;
                ReleaseComObject(usedrange);

                for (int x = 1; x < Xmax; x++)
                {
                    for (int y = 1; y < Ymax; y++)
                    {
                        //Debug.WriteLine(x + " : " + y);
                        string val = GetCellValue(objSheet, x, y) as string;
                        if (!string.IsNullOrEmpty(val))
                        {
                            val = val.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ').Trim();
                            if (val.Contains(sword))
                            {
                                table.AddExcelGrepResultRow(excelFile.FullName, sheetName + ":" + x + "," + y, val);
                                Debug.WriteLine(excelFile.Name + "(" + sheetName + ":" + x + "," + y + ")");
                            }
                        }
                    }
                }
            }

            //Close Book
            if (objBook != null)
            {
                objBook.Close(false, null, null);
                ReleaseComObject(objBook);
            }
            //Close App
            if (objBooks != null)
            {
                objBooks.Close();
                ReleaseComObject(objBooks);
            }
            if (objApp != null)
            {
                objApp.Quit();
                ReleaseComObject(objApp);
            }
        }

        private static void ReleaseComObject(object comObj)
        {
            try
            {
                if (null != comObj) // && Marshal.IsComObject(objCom))
                {
                    int i;
                    do
                    {
                        i = Marshal.ReleaseComObject(comObj);
                    } while (i > 0);
                }
            }
            finally
            {
                comObj = null;
            }
        }

        private static object GetCellValue(Excel.Worksheet worksheet, object rowIndex, object columnIndex)
        {
            Excel.Range range = worksheet.Cells.get_Item(rowIndex, columnIndex) as Excel.Range;
            object obj = range.Value2;
            ReleaseComObject(range);
            return obj;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                if (objApp != null)
                {
                    objApp.Quit();
                }
            }
            finally
            {
                ReleaseComObject(objApp);
            }
        }
    }
}
