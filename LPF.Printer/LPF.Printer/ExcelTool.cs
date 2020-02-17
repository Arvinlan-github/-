using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace LPF.Printer
{
    public class ExcelTool
    {
        #region[office操作]

        #region[变量]
        /// <summary>
        /// Excel操作对象
        /// </summary>
        private Microsoft.Office.Interop.Excel.Application application;
        #endregion

        #region[链接字段]
        /// <summary>
        /// 链接字段
        /// </summary>
        //private string strConnTxt = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Text;FMT=Delimited;HDR=YES;'";
        private const string strConn93 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
        private const string strConn2010 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
        private const string strNonTitleConn93 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=No;IMEX=1;'";
        private const string strNonTitleConn2010 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=No;IMEX=1;'";
        #endregion

        #region[设置链接]
        /// <summary>
        /// 设置链接
        /// </summary>
        /// <param name="path"></param>
        /// <param name="isTitle">第一行是否为列标题：是-true,否-false</param>
        /// <returns></returns>
        private string GetConnString(string path, bool isTitle = true)
        {
            string strConn = string.Empty;
            if (isTitle) //第一行也当成列标题读取
            {
                if (path.Contains(".xlsx"))
                {
                    strConn = string.Format(strConn2010, path);
                }
                else if (path.Contains(".xls"))
                {
                    strConn = string.Format(strConn93, path);
                }
            }
            else   //第一行也当成数据内容读取
            {
                if (path.Contains(".xlsx"))
                {
                    strConn = string.Format(strNonTitleConn2010, path);
                }
                else if (path.Contains(".xls"))
                {
                    strConn = string.Format(strNonTitleConn93, path);
                }
            }

            return strConn;
        }

        private string Get2010ConnString(string path, bool isTitle = true)
        {
            string strConn = string.Empty;
            if (isTitle) //第一行也当成列标题读取
            {
                if (path.Contains(".xlsx") || path.Contains(".xls"))
                {
                    strConn = string.Format(strConn2010, path);
                }
            }
            else   //第一行也当成数据内容读取
            {
                if (path.Contains(".xlsx") || path.Contains(".xls"))
                {
                    strConn = string.Format(strNonTitleConn2010, path);
                }
            }
            return strConn;
        }
        #endregion

        #region[将EXCEL数据读进DATATABLE]
        /// <summary>
        /// 将EXCEL数据读进DATATABLE
        /// </summary>
        /// <param name="path"></param>
        /// <param name="isTitle">第一行是否为列标题：是-true,否-false</param>
        /// <returns></returns>
        public System.Data.DataTable ReadExcelToDataTable(string path, bool isTitle = true, bool isFirstSheet = true)
        {
            System.Data.DataTable dtRst = null;
            try
            {
                OleDbConnection cnnxls = new OleDbConnection(GetConnString(path, isTitle));
                try
                {
                    cnnxls.Open();
                }
                catch (Exception)
                {
                    cnnxls = new OleDbConnection(Get2010ConnString(path, isTitle));
                    cnnxls.Open();
                }
                System.Data.DataTable dtSchema = cnnxls.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dtSchema.Rows.Count == 0) { return null; }
                OleDbDataAdapter oda = new OleDbDataAdapter();
                if (isFirstSheet)
                {
                    string firstSheetName = GetFirstSheetName(path, 1);
                    oda = new OleDbDataAdapter(string.Format("select * from [{0}]",
                 firstSheetName + '$'), cnnxls);
                }
                else
                {
                    oda = new OleDbDataAdapter(string.Format("select * from [{0}]",
                      dtSchema.Rows[0]["TABLE_NAME"].ToString()), cnnxls);
                }
                DataSet ds = new DataSet();
                //将Excel里面有表内容装载到内存表中！
                oda.Fill(ds);
                if (ds.Tables.Count > 0)
                {
                    dtRst = ds.Tables[0];
                }
                cnnxls.Close();
                return dtRst;
            }
            catch
            {
                return null;
            }
        }
        #endregion

        #region [获取EXCEL中的第一个表名]
        public static string GetFirstSheetName(string filepath, int numberSheetID)
        {
            if (!System.IO.File.Exists(filepath))
            {
                return "This file is on the sky??";
            }
            if (numberSheetID <= 1) { numberSheetID = 1; }
            try
            {
                Microsoft.Office.Interop.Excel.Application obj = default(Microsoft.Office.Interop.Excel.Application);
                Microsoft.Office.Interop.Excel.Workbook objWB = default(Microsoft.Office.Interop.Excel.Workbook);
                string strFirstSheetName = null;
                obj = (Microsoft.Office.Interop.Excel.Application)Microsoft.VisualBasic.Interaction.CreateObject("Excel.Application", string.Empty);
                objWB = obj.Workbooks.Open(filepath, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                strFirstSheetName = ((Microsoft.Office.Interop.Excel.Worksheet)objWB.Worksheets[1]).Name;

                objWB.Close(Type.Missing, Type.Missing, Type.Missing);
                objWB = null;
                obj.Quit();
                obj = null;
                return strFirstSheetName;
            }
            catch (Exception Err)
            {
                return Err.Message;
            }
        }
        #endregion

        #region[将DATATABLE数据写入EXCEL]
        /// <summary>
        /// 将DATATABLE数据写入EXCEL
        /// </summary>
        /// <param name="table"></param>
        /// <param name="excelFilePath"></param>
        public void ReadDataTableToExcel(System.Data.DataTable table, string excelFilePath)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            object missing = System.Reflection.Missing.Value;
            try
            {
                if (xlApp == null)
                {
                    return;
                }
                Microsoft.Office.Interop.Excel.Workbooks xlBooks = xlApp.Workbooks;
                Microsoft.Office.Interop.Excel.Workbook xlBook = xlBooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet xlSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlBook.Worksheets[1];
                xlApp.Visible = false;
                object[,] objData = new object[table.Rows.Count + 1, table.Columns.Count];
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    objData[0, i] = table.Columns[i].ColumnName;
                }
                if (table.Rows.Count > 0)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        for (int j = 0; j < table.Columns.Count; j++)
                        {
                            objData[i + 1, j] = table.Rows[i][j];
                        }
                    }
                }
                string startCol = "A";
                int iCnt = (table.Columns.Count / 26);
                string endColSignal = (iCnt == 0 ? "" : ((char)('A' + (iCnt - 1))).ToString());
                string endCol = endColSignal + ((char)('A' + table.Columns.Count - iCnt * 26 - 1)).ToString();
                Microsoft.Office.Interop.Excel.Range range = xlSheet.get_Range(startCol + "1", endCol + (table.Rows.Count + 1).ToString());
                range.Value = objData;
                range.EntireColumn.AutoFit();
                xlApp.DisplayAlerts = false;
                xlApp.AlertBeforeOverwriting = false;
                if (xlSheet != null)
                {
                    xlSheet.SaveAs(excelFilePath, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    Kill(xlApp);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region[结束指定EXCEL进程]
        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int Processid);
        /// <summary>
        /// 结束指定EXCEL进程
        /// </summary>
        /// <param name="theApp"></param>
        private void Kill(Microsoft.Office.Interop.Excel.Application theApp)
        {
            int iId = 0;
            IntPtr intptr = new IntPtr(theApp.Hwnd);
            System.Diagnostics.Process p = null;
            try
            {
                GetWindowThreadProcessId(intptr, out iId);
                p = System.Diagnostics.Process.GetProcessById(iId);
                if (p != null)
                {
                    p.Kill();
                    p.Dispose();
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        #endregion

        #region[获取EXCEL工作簿Workbook]
        /// <summary>
        /// 获取EXCEL工作簿Workbook
        /// </summary>
        /// <returns></returns>
        public Workbook GetWorkbook(string FilePath)
        {
            if (application == null) { application = new Microsoft.Office.Interop.Excel.Application(); }
            application.DisplayAlerts = false;
            application.AlertBeforeOverwriting = false;
            Microsoft.Office.Interop.Excel.Workbook workbook = application.Workbooks.Open(FilePath, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            return workbook;
        }
        #endregion

        #region[把工作卡SHEET复制到最后一个卡并返回]
        /// <summary>
        /// 把工作卡SHEET复制到最后一个卡并返回
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public Worksheet CopyToLastSheet(Workbook workbook, Worksheet sheet)
        {
            Worksheet _worksheet = new Worksheet();
            sheet.Copy(Type.Missing, workbook.Sheets[workbook.Sheets.Count]);
            _worksheet = workbook.Sheets[workbook.Sheets.Count];
            return _worksheet;
        }
        #endregion

        #region [获取指定excel的指定单元格内容]
        ///<summary>
        /// 获取指定excel的指定单元格内容（此种方法获取第一行数据时会异常,暂无法处理）
        ///</summary>
        /// <param name="fileName">文件路径</param>
        /// <param name="row">行号(需要大于0，否则会导致异常)</param>
        /// <param name="column">列号</param>
        /// <returns>返回单元指定单元格内容</returns>
        public string GetExcelOneCell(string fileName, int row, int column)
        {
            Application app = new Application();
            Workbook wbook = app.Workbooks.Open(fileName, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
            Worksheet workSheet = new Worksheet();
            workSheet = (Worksheet)wbook.Worksheets[wbook.Worksheets.Count];
            string temp = ((Range)workSheet.Cells[row, column]).Text.ToString();
            wbook.Close(false, fileName, false);
            app.Quit();
            NAR(app);
            NAR(wbook);
            NAR(workSheet);
            return temp;
        }
        /// <summary>
        /// 此函数用来释放对象的相关资源
        /// </summary>
        /// <param name="o"></param>
        private void NAR(Object o)
        {
            try
            {
                //使用此方法，来释放引用某些资源的基础 COM 对象。 这里的o就是要释放的对象
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch { }
            finally
            {
                o = null; GC.Collect();
            }
        }
        #endregion

        #region[结束EXCEL进程]
        /// <summary>
        /// 结束指定excel进程，如果没指定则结束该操作类的excel操作对象
        /// </summary>
        /// <param name="theApp"></param>
        public void KillEXCEL(Microsoft.Office.Interop.Excel.Application theApp = null)
        {
            if (theApp != null) { Kill(theApp); }
            else if (application != null) { Kill(application); }
        }
        #endregion

        #endregion

        #region[AppLibrary操作]

        /// <summary>
        /// 将DATATABLE数据写入EXCEL
        /// </summary>
        /// <param name="table">操作表</param>
        /// <param name="excelFilePath">excel保存路径</param>
        /// <param name="excelFileName">excel文件名</param>
        public void ReadDataTableToExcelForAppLibrary(System.Data.DataTable table, string excelFilePath, string excelFileName)
        {
            AppLibrary.WriteExcel.XlsDocument doc = new AppLibrary.WriteExcel.XlsDocument();
            doc.FileName = excelFileName;
            string SheetName = "sheet1";
            AppLibrary.WriteExcel.Worksheet sheet = doc.Workbook.Worksheets.Add(SheetName);
            AppLibrary.WriteExcel.Cells cells = sheet.Cells;
            for (int i = 1; i <= table.Columns.Count; i++)
            {
                cells.Add(1, i, table.Columns[i - 1].ColumnName);
            }
            for (int r = 1; r <= table.Rows.Count; r++)
            {
                for (int c = 1; c <= table.Columns.Count; c++)
                {
                    if (table.Columns[c - 1].DataType.Name.ToUpper().Contains("DECIMAL") || table.Columns[c - 1].DataType.Name.ToUpper().Contains("DOUBLE") || table.Columns[c - 1].DataType.Name.ToUpper().Contains("INT"))
                    {
                        try
                        {
                            if (table.Rows[r - 1][c - 1].GetType().Name == "DBNull")
                            {
                                cells.Add(r + 1, c, 0);
                            }
                            else
                            {
                                cells.Add(r + 1, c, table.Rows[r - 1][c - 1]);
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                    else
                    {
                        cells.Add(r + 1, c, table.Rows[r - 1][c - 1].ToString());
                    }
                }
            }
            doc.Save(excelFilePath, true);
        }
        #endregion
    }
}
