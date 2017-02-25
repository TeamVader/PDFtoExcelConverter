﻿using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using PDFtoExcelConverter;
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
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace PDFtoExcelConverter
{
    public partial class MainForm : Form
    {
        private bool File_Selected;
        private string path_to_pdf;
        private Int32 no_copies;
        private string startup_path;
        private Stopwatch sw;
        private BackgroundWorker worker;
        public static int size = 2000;
        private string Exceltemplate = "BMK-Vorlage.xls";
        public MainForm()
        {
            InitializeComponent();

        }

        private void buttopen_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "PDF (*.pdf)|*.pdf";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Multiselect = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                path_to_pdf = openFileDialog1.FileName;
                if (System.IO.Path.GetExtension(path_to_pdf) == ".pdf")
                {
                    File_Selected = true;
                    buttconvert.Enabled = true;
                    labelname.Text = System.IO.Path.GetFileNameWithoutExtension(path_to_pdf);
                }
                else
                {
                    File_Selected = false;
                    buttconvert.Enabled = false;
                    labelname.Text = "No PDF!!";
                }
            }
            else
            {
                //File_Selected = false;
                buttconvert.Enabled = false;
                labelname.Text = "";
            }
        }

        private void cbocopies_SelectedIndexChanged(object sender, EventArgs e)
        {
            // MessageBox.Show(cbocopies.SelectedItem.ToString());
            no_copies = Int32.Parse(cbocopies.SelectedItem.ToString());
        }

        private void buttconvert_Click(object sender, EventArgs e)
        {
            buttconvert.Enabled = false;
            progressBar1.Visible = true;
            worker.RunWorkerAsync();
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            sw = new Stopwatch();
            Regex bmk_regex = new Regex(@"[-+]?([0-9]{1,4})[A-Z][-+]?([0-9]*\.[0-9]{1,3}|[0-9]{1,3})");
            

            int percentFinished = 0;
            //int increment=0;
            // Match match;
            try
            {
                sw.Start();
                Microsoft.Office.Interop.Excel.Application excel_app = new Microsoft.Office.Interop.Excel.Application();
                Range worksheet_range;
                // Make Excel visible (optional).
                excel_app.Visible = false;
                excel_app.DisplayAlerts = false;
                int rowpointer = 0;
                int columnpointer = 0;
                int minrow = 1;
                int mincolumn = 1;
                int maxrow = 44;
                int maxcolumn = 21;
                int sheet_number = 1;
                string[] resultstring;
                string[] lookup = new string[size];
                string[] clamp_lookup = new string[size];
                string[] cable_lookup = new string[size];
                int arraypointer = 0;
                bool saved = false;
                var text = new StringBuilder();

                // The PdfReader object implements IDisposable.Dispose, so you can
                // wrap it in the using keyword to automatically dispose of it
                using (var pdfReader = new PdfReader(path_to_pdf))
                {
                    worker.ReportProgress(percentFinished);

                    // Loop through each page of the document
                    for (var page = 1; page <= pdfReader.NumberOfPages; page++)
                    {

                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();

                        var currentText = PdfTextExtractor.GetTextFromPage(
                            pdfReader,
                            page,
                            strategy);

                        currentText =
                            Encoding.UTF8.GetString(Encoding.Convert(
                                Encoding.Default,
                                Encoding.UTF8,
                                Encoding.Default.GetBytes(currentText)));

                        text.Append(currentText);
                        percentFinished = (page / pdfReader.NumberOfPages) * 100;
                        worker.ReportProgress(percentFinished);

                    }


                    // MessageBox.Show(startup_path + "\\" + Exceltemplate);
                    if (File.Exists(startup_path + "\\" + Exceltemplate))
                    {
                        Microsoft.Office.Interop.Excel.Workbooks workbooks = excel_app.Workbooks;

                        Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Open(
                        (startup_path + "\\" + Exceltemplate));

                        // Get the first worksheet.
                        Microsoft.Office.Interop.Excel.Sheets sheets = workbook.Sheets;
                        Microsoft.Office.Interop.Excel.Worksheet sheet_template = sheets["Tabelle1"];

                        rowpointer = minrow;
                        columnpointer = mincolumn;

                        if (sheet_template != null)
                        {
                            resultstring = text.ToString().Split(new char[] { ' ' });
                            Sort_and_Filter_Algorithms.Filter_Clamp_BMK(clamp_lookup, resultstring);
                            Sort_and_Filter_Algorithms.Filter_Cable_BMK(cable_lookup, resultstring);

                            for (int j = 0; j < resultstring.Length; j++)
                            {
                                
                                //
                                foreach (Match match in bmk_regex.Matches(resultstring[j]))
                                {



                                    if (findstring(match.Value.Replace("-", ""), lookup) == false)
                                    {

                                        lookup[arraypointer] = match.Value.Replace("-", "");
                                        arraypointer++;

                                    }
                                }

                            }
                            
                            Array.Sort(lookup, StringComparer.InvariantCulture);
                           
                            for (int i = 0; i < lookup.Length; i++)
                            {

                                if (columnpointer > maxcolumn)
                                {
                                    columnpointer = mincolumn;
                                    rowpointer += no_copies;
                                }

                                if ((no_copies + rowpointer) > (maxrow + 1))
                                {
                                    workbook.SaveAs(System.IO.Path.GetDirectoryName(path_to_pdf) + "\\" + System.IO.Path.GetFileNameWithoutExtension(path_to_pdf) + "_" + sheet_number.ToString() + ".xls"); //, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                                    for (int l = minrow; l == maxrow; l++)
                                    {
                                        for (int n = mincolumn; n == maxcolumn; n++)
                                        {
                                            sheet_template.Cells[l, n] = "";
                                        }
                                    }
                                    sheet_number++;
                                    rowpointer = minrow;
                                    saved = true;
                                }

                                if (!string.IsNullOrEmpty(lookup[i]))
                                {
                                    for (int k = 0; k < no_copies; k++)
                                    {
                                        saved = false;
                                        sheet_template.Cells[rowpointer + k, columnpointer] = lookup[i];
                                        //  MessageBox.Show(columnpointer.ToString() + " Reihe" + rowpointer.ToString());



                                    }

                                    columnpointer += 2;
                                }


                            }

                        }
                        if (saved == false)
                        {
                            workbook.SaveAs(System.IO.Path.GetDirectoryName(path_to_pdf) + "\\" + System.IO.Path.GetFileNameWithoutExtension(path_to_pdf) + "_" + sheet_number.ToString() + ".xls"); //, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                        }

                        
                        // workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, excel_path[0] + "_bom.pdf");
                        // Close the workbook without saving changes.
                        XML_Functions.Create_Kabel_XML_File(System.IO.Path.GetDirectoryName(path_to_pdf) + "\\" + "Kabelbeschriftungen.wscx", cable_lookup, no_copies);
                        XML_Functions.Create_Klemme_XML_File(System.IO.Path.GetDirectoryName(path_to_pdf) + "\\" + "Klemmenbeschriftungen.wscx", clamp_lookup, no_copies);

                        lookup = null;
                        resultstring = null;
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet_template);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets);
                        workbook.Close(0);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                        excel_app.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel_app);

                        foreach (Process process in Process.GetProcessesByName("Excel"))
                        {
                            if (!string.IsNullOrEmpty(process.ProcessName) && process.StartTime.AddSeconds(+10) > DateTime.Now)
                            {
                                process.Kill();
                            }
                        }

                        Process.Start("explorer.exe", System.IO.Path.GetDirectoryName(path_to_pdf));
                    }
                    else
                    {
                        // labelname.Text = "No Template in Folder";
                    }
                    //MessageBox.Show(text.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }






            // System.Threading.Thread.Sleep(50);

        }

        public class CustomComparer : IComparer<string>
        {
            public int Compare(string x, string y)
            {
                var regex = new Regex(@"[-+]?([0-9])");

                // run the regex on both strings
                var xRegexResult = regex.Match(x);
                var yRegexResult = regex.Match(y);

                // check if they are both numbers
                if (xRegexResult.Success && yRegexResult.Success)
                {
                    return int.Parse(xRegexResult.Groups[1].Value).CompareTo(int.Parse(yRegexResult.Groups[1].Value));
                }

                // otherwise return as string comparison
                return x.CompareTo(y);
            }
        }


        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 0)
            {
                labelname.Text = "Converting File";
            }

            progressBar1.Value = e.ProgressPercentage;
        }


        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Visible = false;
            buttconvert.Enabled = true;
            sw.Stop();
            labelstopwatch.Text = "Benchmark :" + sw.Elapsed.Milliseconds + "ms";
            labelname.Text = "File successfully converted";
        }

        /// <summary>
        /// Lookup for String in array
        /// </summary>
        /// <param name="search"></param>
        /// <param name="array"></param>
        /// <returns></returns>
        private static bool findstring(string search, string[] array)
        {
            for (int i = 0; i < array.Length; i++)
            {
                if (search == array[i])
                {
                    return true;
                }
            }
            return false;
        }

       

        private void Form1_Load(object sender, EventArgs e)
        {
            File_Selected = false;
            path_to_pdf = "";
            buttconvert.Enabled = false;
            cbocopies.Items.AddRange(new string[] { "1", "2", "3", "4", "5" });
            no_copies = 1;
            cbocopies.SelectedIndex = 0;
            startup_path = (System.Windows.Forms.Application.StartupPath);
            // MessageBox.Show(startup_path);
            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;

            worker.DoWork += new DoWorkEventHandler(worker_DoWork);
            worker.ProgressChanged +=
                        new ProgressChangedEventHandler(worker_ProgressChanged);
            worker.RunWorkerCompleted +=
                       new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);
        }



    }
}