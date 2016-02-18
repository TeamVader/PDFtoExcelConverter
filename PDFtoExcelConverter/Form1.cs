using iTextSharp.text.pdf;
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
    public partial class Form1 : Form
    {
        private bool File_Selected;
        private string path_to_pdf;
        private Int32 no_copies;
        private string startup_path;

        private string Exceltemplate = "BMK-Vorlage.xls";
        public Form1()
        {
            InitializeComponent();
            File_Selected = false;
            path_to_pdf = "";
            buttconvert.Enabled = false;
            cbocopies.Items.AddRange(new string[] { "1", "2" ,"3","4","5"});
            no_copies = 1;
            cbocopies.SelectedIndex = 0;
            startup_path = (System.Windows.Forms.Application.StartupPath);
           // MessageBox.Show(startup_path);
            
        }

        private void buttopen_Click(object sender, EventArgs e)
        {
           
         OpenFileDialog openFileDialog1 = new OpenFileDialog();

         openFileDialog1.InitialDirectory = "c:\\" ;
          openFileDialog1.Filter = "PDF (*.pdf)|*.pdf" ;
          openFileDialog1.FilterIndex = 2 ;
          openFileDialog1.RestoreDirectory = true ;
          openFileDialog1.Multiselect = false;

          if (openFileDialog1.ShowDialog() == DialogResult.OK)
          {
              path_to_pdf = openFileDialog1.FileName;
              if (System.IO.Path.GetExtension(path_to_pdf) == ".pdf")
              {
                  File_Selected=true;
                  buttconvert.Enabled=true;
                  labelname.Text = System.IO.Path.GetFileNameWithoutExtension(path_to_pdf);
              }
              else
              {
                  File_Selected = false;
                      buttconvert.Enabled=false;
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

            Regex bmk_regex = new Regex(@"[-+]?([0-9]*\.[0-9]+|[0-9]+)[A-Z][-+]?([0-9]*\.[0-9]+|[0-9]+)");
           // Match match;
            try
            {

                 Microsoft.Office.Interop.Excel.Application excel_app = new Microsoft.Office.Interop.Excel.Application();
                 Range worksheet_range;
                // Make Excel visible (optional).
                 excel_app.Visible = false;
                 excel_app.DisplayAlerts = false;
                int rowpointer=0;
                int columnpointer =0;
                int minrow = 2;
                int mincolumn = 3;
                int maxrow=43;
                int maxcolumn=21;
                int sheet_number = 1;
                string[] resultstring;
                string[] lookup = new string[2000];
                int arraypointer=0;
                bool saved = false;
                var text = new StringBuilder();

                // The PdfReader object implements IDisposable.Dispose, so you can
                // wrap it in the using keyword to automatically dispose of it
                using (var pdfReader = new PdfReader(path_to_pdf))
                {
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
                    }

                   // MessageBox.Show(startup_path + "\\" + Exceltemplate);
                    if(File.Exists(startup_path+"\\"+Exceltemplate))
                    {
                    Microsoft.Office.Interop.Excel.Workbooks workbooks = excel_app.Workbooks;

                    Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Open(
                    (startup_path+"\\"+Exceltemplate));

                    // Get the first worksheet.
                    Microsoft.Office.Interop.Excel.Sheets sheets = workbook.Sheets;
                    Microsoft.Office.Interop.Excel.Worksheet sheet_template = sheets["Vorlage"];

                    rowpointer = minrow;
                    columnpointer = mincolumn;
                    
                    if (sheet_template != null)
                    {
                        resultstring = text.ToString().Split(new char[]{' '});
                        for (int i = 0; i < resultstring.Length; i++)
                        {
                          

                            //
                            foreach(Match match in bmk_regex.Matches(resultstring[i]))
                            {

                                if (columnpointer > maxcolumn)
                                {
                                    columnpointer = mincolumn;
                                    rowpointer += no_copies;
                                }

                                if ((no_copies + rowpointer) > (maxrow+1))
                                {
                                    workbook.SaveAs(System.IO.Path.GetDirectoryName(path_to_pdf) + "\\" + System.IO.Path.GetFileNameWithoutExtension(path_to_pdf) + "_" + sheet_number.ToString() + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
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

                                if (findstring(match.Value, lookup) == false)
                                {
                                    for (int k = 0; k < no_copies; k++)
                                    {
                                        saved = false;
                                        sheet_template.Cells[rowpointer + k, columnpointer] = match.Value;
                                        //  MessageBox.Show(columnpointer.ToString() + " Reihe" + rowpointer.ToString());



                                    }
                                    lookup[arraypointer] = match.Value;
                                    arraypointer++;
                                    columnpointer += 2;
                                }
                            }
                            
                           
                        }
                       
                    }
                    if (saved==false)
                    {
                        workbook.SaveAs(System.IO.Path.GetDirectoryName(path_to_pdf) + "\\" + System.IO.Path.GetFileNameWithoutExtension(path_to_pdf) + "_" + sheet_number.ToString() + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                    }
                            lookup=null;
                    resultstring=null;

                    // workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, excel_path[0] + "_bom.pdf");
                    // Close the workbook without saving changes.
                    labelname.Text = "File successfully converted";
                   
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
                    }
                    else
                    {
                        labelname.Text="No Template in Folder";
                    }
                    //MessageBox.Show(text.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        /// <summary>
        /// Lookup for String in array
        /// </summary>
        /// <param name="search"></param>
        /// <param name="array"></param>
        /// <returns></returns>
        private static bool findstring(string search,string[] array)
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

        }
        

        
    }
}
