using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LPF.Printer
{
    public partial class FrmSetRule : Form
    {
        public CodeRule _coderule = new CodeRule();

        public FrmSetRule()
        {
            InitializeComponent();
        }

        private void FrmSetRule_Load(object sender, EventArgs e)
        {

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            _coderule.lenght = string.IsNullOrEmpty(txtLenght.Text) ? "0" : txtLenght.Text;
            _coderule.startnumber = string.IsNullOrEmpty(txtStartNumber.Text) ? "0" : txtStartNumber.Text;
            _coderule.endnumber = string.IsNullOrEmpty(txtEnd.Text) ? "0" : txtEnd.Text;
            _coderule.display = txtDisplay.Text;
        }

        private void txtLenght_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtLenght.Text.Trim())) { return; }
            int i = txtTryParse(txtLenght.Text.Trim());
            if (i == -1)
            {
                txtStartNumber.Text = "";
            }
            else
            {
                txtDisplay.Text = string.IsNullOrEmpty(txtStartNumber.Text) ? 0.ToString().PadLeft(i, '0') : txtStartNumber.Text.PadLeft(i, '0');
            }
        }

        private void txtStartNumber_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtStartNumber.Text.Trim())) { return; }
            int i = txtTryParse(txtStartNumber.Text.Trim());
            if (i == -1)
            {
                txtStartNumber.Text = "";
            }
            else
            {
                txtDisplay.Text = string.IsNullOrEmpty(txtLenght.Text) ? "" : txtStartNumber.Text.PadLeft(Convert.ToInt32(txtLenght.Text), '0');
            }
        }

        private int txtTryParse(string value)
        {
            int number = -1;
            bool temp = int.TryParse(value, out number);
            if (temp)
            {
                return number;
            }
            else
            {
                MessageBox.Show("请输入有效的数字", "提示");
                return -1;
            }
        }

    }
}

