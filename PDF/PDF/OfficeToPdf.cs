using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace PDF
{
    public class OfficeToPdf
    {
        /// <summary>
        /// 把Word或文本文件转换成为PDF格式文件
        /// </summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param> 
        /// <returns>true=转换成功</returns>
        public static bool DOCConvertToPDF(string sourcePath, string targetPath)
        {
            bool result = false;
            Word.WdExportFormat exportFormat = Word.WdExportFormat.wdExportFormatPDF;
            object paramMissing = Type.Missing;
            Word.ApplicationClass wordApplication = new Word.ApplicationClass();
            Word.Document wordDocument = null;
            try
            {
                object paramSourceDocPath = sourcePath;
                string paramExportFilePath = targetPath;

                Word.WdExportFormat paramExportFormat = exportFormat;
                bool paramOpenAfterExport = false;
                Word.WdExportOptimizeFor paramExportOptimizeFor = Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
                Word.WdExportRange paramExportRange = Word.WdExportRange.wdExportAllDocument;
                int paramStartPage = 0;
                int paramEndPage = 0;
                Word.WdExportItem paramExportItem = Word.WdExportItem.wdExportDocumentContent;
                bool paramIncludeDocProps = true;
                bool paramKeepIRM = true;
                Word.WdExportCreateBookmarks paramCreateBookmarks = Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                bool paramDocStructureTags = true;
                bool paramBitmapMissingFonts = true;
                bool paramUseISO19005_1 = false;
                //Logger.WriteLog("word开始", "log/Catalog_Info.txt", true);
                wordDocument = wordApplication.Documents.Open(
                ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                ref paramMissing, ref paramMissing, ref paramMissing,
                ref paramMissing, ref paramMissing, ref paramMissing,
                ref paramMissing, ref paramMissing, ref paramMissing,
                ref paramMissing, ref paramMissing, ref paramMissing,
                ref paramMissing);

                if (wordDocument != null)
                    wordDocument.ExportAsFixedFormat(paramExportFilePath,
                    paramExportFormat, paramOpenAfterExport,
                    paramExportOptimizeFor, paramExportRange, paramStartPage,
                    paramEndPage, paramExportItem, paramIncludeDocProps,
                    paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                    paramBitmapMissingFonts, paramUseISO19005_1,
                    ref paramMissing);
                result = true;
            }
            catch (Exception ex)
            {
                result = false;
                //Logger.WriteLog("Service Error:Word转pdf" + ex.ToString(), "log/Catalog_Info.txt", true);

            }
            finally
            {
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordDocument = null;
                }
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordApplication = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }
        /// <summary>
        /// 把Excel文件转换成PDF格式文件
        /// </summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param> 
        /// <returns>true=转换成功</returns>
        public static bool XLSConvertToPDF(string sourcePath, string targetPath)
        {

            // ConvertExcelPDF(sourcePath, targetPath);
            bool result = false;
            Excel.XlFixedFormatType targetType = Excel.XlFixedFormatType.xlTypePDF;
            object missing = Type.Missing;
            Excel.ApplicationClass application = null;
            Excel.Workbook workBook = null;

            try
            {
                application = new Excel.ApplicationClass();
                object target = targetPath;
                object type = targetType;

                workBook = application.Workbooks.Open(sourcePath, missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing, missing, missing, missing, missing);

                #region 设置打印参数
                foreach (Microsoft.Office.Interop.Excel.Worksheet sh in workBook.Sheets)
                {
                    sh.PageSetup.Zoom = false;
                    sh.PageSetup.FitToPagesTall = 1;
                    sh.PageSetup.FitToPagesWide = 1;
                    sh.PageSetup.TopMargin = 0;
                    sh.PageSetup.BottomMargin = 0;
                    sh.PageSetup.LeftMargin = 0;
                    sh.PageSetup.RightMargin = 0;
                    sh.PageSetup.CenterHorizontally = true;
                }
                #endregion


                workBook.ExportAsFixedFormat(targetType, target, Excel.XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);

                result = true;
            }
            catch (Exception ex)
            {
                result = false;
                //Logger.WriteLog("Service Error:Excel转pdf" + ex.ToString(), "Log\\ServiceInfoError.txt", true);
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(true, missing, missing);
                    workBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        /// <summary>
        /// 转换excel为pdf
        /// </summary>
        /// <param name="filename">doc文件路径</param>
        /// <param name="savefilename">pdf保存路径</param>
        public static void ConvertExcelPDF(string filename, string PDFFileName)
        {
            //先引入：Microsoft.Office.Interop 
            //再 using Microsoft.Office.Interop.Excel; 


            Microsoft.Office.Interop.Excel.ApplicationClass cExcel = new Microsoft.Office.Interop.Excel.ApplicationClass();
            cExcel.Visible = true;
            object missing = Type.Missing;
            //excel 文件名：
            string excelFileName = filename.ToString();

            Int32 intLastDot = excelFileName.LastIndexOf(".");

            string tmp_FilePath = excelFileName.Substring(0, intLastDot);


            Microsoft.Office.Interop.Excel.Workbook book = null;

            try
            {
                book = cExcel.Workbooks.Open(excelFileName, missing, missing, missing, missing
                     , missing, missing, missing, missing, missing, missing
                     , missing, missing, missing, missing);

                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Worksheets[1];

                //参数1
                var formatType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
                //参数3
                var quarlity = Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard;

                //倒数第二个参数控制是否用pdf软件打开，false会在后台处理，不打开文件
                sheet.ExportAsFixedFormat(formatType, PDFFileName, quarlity, true, true, missing, missing, false, missing);

            }
            catch (Exception ex)
            {

            }
            finally
            {
                if (book != null)
                {
                    book.Close(missing, missing, missing);
                    book = null;
                }
                cExcel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(cExcel);
                cExcel = null;
                GC.Collect();
            }



        }




        /// <summary>
        /// 把PowerPoing文件转换成PDF格式文件
        /// </summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param> 
        /// <returns>true=转换成功</returns>
        public static bool PPTConvertToPDF(string sourcePath, string targetPath)
        {

            bool result;
            PowerPoint.PpSaveAsFileType targetFileType = PowerPoint.PpSaveAsFileType.ppSaveAsPDF;
            object missing = Type.Missing;
            PowerPoint.ApplicationClass application = null;
            PowerPoint.Presentation persentation = null;
            try
            {
                application = new PowerPoint.ApplicationClass();
                persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                persentation.SaveAs(targetPath, targetFileType, Microsoft.Office.Core.MsoTriState.msoTrue);

                result = true;
            }
            catch (Exception ex)
            {
                result = false;
                //Logger.WriteLog("Service Error:PowerPoing转pdf" + ex.ToString(), "Log\\ServiceInfoError.txt", true);
            }
            finally
            {
                if (persentation != null)
                {
                    persentation.Close();
                    persentation = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        /// <summary>
        /// 导出数据到pdf文件
        /// </summary>
        /// <param name="dt">要导出的数据表</param>
        /// <param name="path">文件路径</param>
        public static void ExportToPDF(System.Data.DataTable dt, string path)
        {
            try
            {
                if (dt.Columns.Count == 1)
                    dt.Columns.Add("  ");
                float wh = dt.Columns.Count * 72.0f;
                iTextSharp.text.Rectangle rec = new iTextSharp.text.Rectangle(0.0f, 0.0f, wh, wh);
                iTextSharp.text.Document document = new iTextSharp.text.Document(rec, 36.0f, 36.0f, 36.0f, 36.0f);
                System.IO.FileStream fs = new System.IO.FileStream(path, FileMode.Create);
                PdfWriter.GetInstance(document, fs); //在当前路径下创一个文件 　
                document.Open();
                BaseFont bfChinese = null;
                if (File.Exists(@"C:\Windows\Fonts\simsun.ttf"))
                    bfChinese = BaseFont.CreateFont(@"C:\Windows\Fonts\simsun.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                else if (File.Exists(@"simsun.ttf"))
                    bfChinese = BaseFont.CreateFont(@"simsun.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                iTextSharp.text.Font fontChinese = new iTextSharp.text.Font(bfChinese, 10, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                PdfPTable table = new PdfPTable(dt.Columns.Count);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        table.AddCell(new Phrase(dt.Rows[i][j].ToString(), fontChinese));
                    }

                }

                try
                {
                    //表格相对空白宽度占比
                    table.WidthPercentage = 100.0f;
                    document.Add(table);　　//添加table
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                document.Close();
                document.Dispose();

            }
            catch (DocumentException ex)
            {
                throw ex;
            }
        }

        public static DataSet GetExcel(string fileName)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Workbook oWB;
            Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            try
            {
                //  creat a Application object
                oXL = new Microsoft.Office.Interop.Excel.ApplicationClass();
                //   get   WorkBook  object
                oWB = oXL.Workbooks.Open(fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Missing.Value);

                //   get   WorkSheet object 
                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.Sheets[1];
                System.Data.DataTable dt = new System.Data.DataTable("dtExcel");
                DataSet ds = new DataSet();
                ds.Tables.Add(dt);
                DataRow dr;

                StringBuilder sb = new StringBuilder();
                int jValue = oSheet.UsedRange.Cells.Columns.Count;
                int iValue = oSheet.UsedRange.Cells.Rows.Count;
                //  get data columns
                for (int j = 1; j <= jValue; j++)
                {
                    dt.Columns.Add("column" + j, System.Type.GetType("System.String"));
                }

                //string colString = sb.ToString().Trim();
                //string[] colArray = colString.Split(':');

                //  get data in cell
                for (int i = 1; i <= iValue; i++)
                {
                    dr = ds.Tables["dtExcel"].NewRow();
                    for (int j = 1; j <= jValue; j++)
                    {
                        oRng = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[i, j];
                        string strValue = oRng.Text.ToString();
                        dr["column" + j] = strValue;
                    }
                    ds.Tables["dtExcel"].Rows.Add(dr);
                }
                return ds;
            }
            catch (Exception ex)
            {

                return null;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
