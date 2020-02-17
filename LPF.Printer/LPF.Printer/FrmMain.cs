using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Reflection;
using System.Linq;

namespace LPF.Printer
{
    public partial class FrmMain : Form
    {
        #region 局部变量
        ProductEntity _productentity = new ProductEntity();
        CodeRule _coderule;
        ExcelTool excel = new ExcelTool();
        #endregion

        #region 构造函数
        public FrmMain()
        {
            InitializeComponent();
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            AddPrintersNameToList();
            InitDgvList();
            dgvList.AllowUserToAddRows = false;
            dgvList.AutoGenerateColumns = false;
        }
        #endregion

        #region 按钮事件
        /// <summary>
        /// 设置流水内容按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSet_Click(object sender, EventArgs e)
        {
            FrmSetRule fsr = new FrmSetRule();
            if (fsr.ShowDialog() == DialogResult.OK)
            {
                _coderule = fsr._coderule;
            }
        }

        private void btnBuildList_Click(object sender, EventArgs e)
        {
            if (_coderule == null || (_coderule.lenght == "0" && _coderule.startnumber == "0" && _coderule.endnumber == "0"))
            {
                MessageBox.Show("请先设置条码生成规则。", "提示");
                return;
            }
            else
            {
                SetProductEntity();
                BuildData();
                //Math.sp
            }
        }
        #endregion

        #region 自定义函数
        /// <summary>
        /// 设置打印对象数据
        /// </summary>
        void SetProductEntity()
        {
            _productentity.productname = txtProductName.Text.Trim();
            _productentity.thick = txtThick.Text.Trim();
            _productentity.weight = txtWeight.Text.Trim();
            _productentity.spec = txtSpec.Text.Trim();
            _productentity.producer = txtProducer.Text.Trim();
            _productentity.productdate = txtProductDate.Text.Trim();
            _productentity.quantity = txtQuantity.Text.Trim();
            _productentity.gram = txtGram.Text.Trim();
            _productentity.level = txtLevel.Text.Trim();
            _productentity.qc = txtQC.Text.Trim();
            _productentity.customercode = txtCustomerCode.Text.Trim();
            _productentity.ordernumber = txtOrderNumber.Text.Trim();
            _productentity.coderule = txtCodeRule.Text.Trim();
        }

        /// <summary>
        /// 设置打印机列表
        /// </summary>
        private void AddPrintersNameToList()
        {
            PrintDocument pd = new PrintDocument();
            lblPrinterName.Text = pd.PrinterSettings.PrinterName;
            List<string> rst = new List<string>();
            foreach (string s in PrinterSettings.InstalledPrinters)
            {
                lvPrinter.Items.Add(s.ToString());
            }
        }

        /// <summary>
        /// 初始化明细列表
        /// </summary>
        void InitDgvList()
        {
            DataGridViewTextBoxColumn dgvCol = new DataGridViewTextBoxColumn();
            dgvCol.HeaderText = "序号";
            dgvCol.Name = "XH";
            dgvCol.DataPropertyName = "XH";
            dgvList.Columns.Add(dgvCol);
            DataGridViewTextBoxColumn dgvCol1 = new DataGridViewTextBoxColumn();
            dgvCol1.HeaderText = "条码";
            dgvCol1.Name = "TM";
            dgvCol1.DataPropertyName = "TM";
            dgvCol1.Width = 200;
            dgvList.Columns.Add(dgvCol1);
            DataGridViewTextBoxColumn dgvCol2 = new DataGridViewTextBoxColumn();
            dgvCol2.HeaderText = "二维码";
            dgvCol2.Name = "EWM";
            dgvCol2.DataPropertyName = "EWM";
            dgvCol2.Width = 400;
            dgvList.Columns.Add(dgvCol2);
        }

        /// <summary>
        /// 创建明细列表数据
        /// </summary>
        void BuildData()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("XH");
            dt.Columns.Add("TM");
            dt.Columns.Add("EWM");
            int index = 0;
            for (int i = Convert.ToInt32(_coderule.startnumber); i <= Convert.ToInt32(_coderule.endnumber); i++)
            {
                index++;
                StringBuilder strB = new StringBuilder();
                strB.AppendFormat("{0}{1}", _productentity.coderule, i.ToString(string.Format("D{0}", _coderule.lenght)));
                #region 二维码内容
                StringBuilder strQR = new StringBuilder();
                strQR.AppendFormat("品名:{0}", _productentity.productname);
                strQR.AppendLine();
                strQR.AppendFormat("caliper:{0}", _productentity.thick);
                strQR.AppendLine();
                strQR.AppendFormat("basis weight:{0}", _productentity.weight);
                strQR.AppendLine();
                strQR.AppendFormat("siz:{0}", _productentity.spec);
                strQR.AppendLine();
                strQR.AppendFormat("No:{0}", strB);
                strQR.AppendLine();
                strQR.AppendFormat("Sheets:{0}", _productentity.quantity);
                strQR.AppendLine();
                strQR.AppendFormat("PO:{0}", _productentity.ordernumber);
                strQR.AppendLine();
                strQR.AppendFormat("{0},{1},{2},{3},{4}", _productentity.customercode, _productentity.level, _productentity.qc, _productentity.productdate, _productentity.producer);
                #endregion
                DataRow dr = dt.NewRow();
                dr["XH"] = index;
                dr["TM"] = strB;
                dr["EWM"] = strQR;
                dt.Rows.Add(dr);
            }
            dgvList.DataSource = dt;
        }

        /// <summary>
        /// 设置一维码
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        bool SetBarcodeInfo(string filename, Worksheet _worksheet, string excelrange, string value)
        {
            bool temp = false;
            Range range = _worksheet.Range[excelrange].MergeArea;
            BarcodeInfo bc = new BarcodeInfo();
            if (bc.CreateBarCode(value, filename))
            {
                _worksheet.Shapes.AddPicture(filename, MsoTriState.msoTrue, MsoTriState.msoTrue, range.Left, range.Top, range.Width, range.Height);
                temp = true;
            }
            else
            {
                MessageBox.Show("条码生成失败！", "标签打印提示");
            }
            return temp;
        }

        /// <summary>
        /// 设置二维码
        /// </summary>
        /// <param name="filename">文件名</param>
        /// <param name="excelrange">excel填充区域</param>
        /// <param name="value">值</param>
        /// <returns></returns>
        bool SetQRCode(string filename, Worksheet _worksheet, string excelrange, string value)
        {
            bool temp = false;
            Range range = _worksheet.Range[excelrange].MergeArea;
            QRCode qr = new QRCode();
            if (qr.CreateQRCode(value, filename))
            {
                _worksheet.Shapes.AddPicture(filename, MsoTriState.msoTrue, MsoTriState.msoTrue, range.Left, range.Top, range.Width, range.Height);
                temp = true;
            }
            else
            {
                MessageBox.Show("条码生成失败！", "标签打印提示");
            }
            return temp;
        }

        void PrintLabel(string barcodevalue, string qrcodevalue)
        {
            try
            {
                string qrcodefile_weight = Path.Combine(Directory.GetCurrentDirectory(), "克重.bmp");
                string qrcodefile = Path.Combine(Directory.GetCurrentDirectory(), "二维.bmp");
                string barcodefile = Path.Combine(Directory.GetCurrentDirectory(), "一维.bmp");
                string printFilePath = Path.Combine(Directory.GetCurrentDirectory(), "标签模板.xls");

                Workbook workbook = excel.GetWorkbook(printFilePath);
                Worksheet sheet = workbook.Sheets["Sheet1"];
                sheet.get_Range("J27").Value = _productentity.productname;
                sheet.get_Range("C35").Value = _productentity.thick;
                sheet.get_Range("I35").Value = _productentity.gram;
                sheet.get_Range("Q35").Value = _productentity.spec;
                sheet.get_Range("AB35").Value = _productentity.quantity;
                sheet.get_Range("H46").Value = _productentity.qc;
                sheet.get_Range("O47").Value = Convert.ToDateTime(_productentity.productdate).Year;
                sheet.get_Range("S47").Value = Convert.ToDateTime(_productentity.productdate).Month;
                sheet.get_Range("V47").Value = Convert.ToDateTime(_productentity.productdate).Day;
                sheet.get_Range("AB47").Value = _productentity.level;
                SetQRCode(qrcodefile_weight, sheet, "B46", _productentity.gram);
                if (chkOneDimensionalCode.Checked)
                {
                    SetBarcodeInfo(barcodefile, sheet, "M20", barcodevalue);
                    sheet.get_Range("M25").Value = barcodevalue;
                }
                else
                {
                    SetBarcodeInfo(barcodefile, sheet, "M20", "");//需要移除图片
                    sheet.get_Range("M25").Value = "";//需要移除图片
                }
                if (chkQRCode.Checked)
                {
                    SetQRCode(qrcodefile, sheet, "E13", qrcodevalue);
                }
                else
                {
                    SetQRCode(qrcodefile, sheet, "E13", "");
                }
                sheet.PrintOutEx(Missing.Value, Missing.Value, Convert.ToInt32(txtPrintQuantity.Text.Trim()), false, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            }
            catch (Exception e)
            {
                MessageBox.Show("标签打印异常", "提示");
            }
        }
        #endregion

        #region 右键功能

        /// <summary>
        /// 设置打印设备按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tsmSetPrinter_Click(object sender, EventArgs e)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrinterSettings.PrinterName = lvPrinter.SelectedItems[0].Text;
            lblPrinterName.Text = pd.PrinterSettings.PrinterName;
        }

        /// <summary>
        /// 列表打印全部按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tsmPrintAll_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = (System.Data.DataTable)dgvList.DataSource;
            foreach (DataRow dr in dt.Rows)
            {
                PrintLabel(dr["TM"].ToString(), dr["EWM"].ToString());
            }
        }

        /// <summary>
        /// 打印选中的数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tsmPrintSelected_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgvList.SelectedRows)
            {
                PrintLabel(row.Cells[1].Value.ToString(), row.Cells[2].Value.ToString());
            }
        }

        /// <summary>
        /// 导出数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tsmOutPutData_Click(object sender, EventArgs e)
        {
            if (dgvList.DataSource == null) { MessageBox.Show("没有需要导出的信息。"); return; }
            OutPut();
        }
        #endregion

        private void txtPrintQuantity_TextChanged(object sender, EventArgs e)
        {
            int number = -1;
            bool temp = int.TryParse(txtPrintQuantity.Text, out number);
            if (!temp)
            {
                MessageBox.Show("请输入有效的数字", "提示");
                txtPrintQuantity.Text = "";
            }
        }

        #region[导出操作]
        /// <summary>
        /// 导出
        /// </summary>
        public void OutPut()
        {
            try
            {
                saveFDL.OverwritePrompt = true;
                saveFDL.DefaultExt = "xls";
                saveFDL.AddExtension = true;
                List<DataGridViewColumn> lstColumn = dgvList.Columns.Cast<DataGridViewColumn>().OrderBy(o => o.DisplayIndex).ToList();
                //lstColumn.RemoveAll(r => r.Visible = false);
                Dictionary<string, string> dicColumnName = new Dictionary<string, string>();
                foreach (DataGridViewColumn col in lstColumn)
                {
                    //if (col.Visible == true && (int)col.Name[0] > 127)//因为数据源查询出来的列名都是汉字
                    if (col.Visible == true)
                    {
                        dicColumnName.Add(col.Name, col.HeaderText);
                    }
                }
                System.Data.DataTable dt = ((System.Data.DataTable)dgvList.DataSource).DefaultView.ToTable(false, dicColumnName.Select(s => s.Key).ToArray());
                foreach (DataColumn col in dt.Columns)
                {
                    col.ColumnName = dicColumnName.First(f => f.Key == col.ColumnName).Value;
                }
                if (saveFDL.ShowDialog() == DialogResult.OK)
                {
                    string strPath = saveFDL.FileName.Substring(0, saveFDL.FileName.LastIndexOf("\\"));
                    string strName = saveFDL.FileName.Substring(saveFDL.FileName.LastIndexOf("\\") + 1);
                    //new ExcelDBHelper().ReadDataTableToExcel(dt, saveFDL.FileName);
                    this.Cursor = Cursors.WaitCursor;
                    new ExcelTool().ReadDataTableToExcelForAppLibrary(dt, strPath, strName);
                    this.Cursor = Cursors.Default;
                    MessageBox.Show("导出成功！", "提示");
                }
            }
            catch (Exception e)
            {
                StringBuilder strB = new StringBuilder();
                strB.AppendLine("导出操作失败");
                strB.Append(e.Message);
                MessageBox.Show(strB.ToString(), "提示");
            }
        }
        #endregion
    }
}
