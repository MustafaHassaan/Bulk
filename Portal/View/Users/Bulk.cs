using Controler;
using Model;
using Portal.View.Users;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Portal
{
    public partial class Employees : Form
    {
        int Charcount = 0;
        InputLanguage MCL = InputLanguage.CurrentInputLanguage;
        public static string UNPassing;
        public static string IDPassing;
        public Employees()
        {
            InitializeComponent();
            LblName.Text = EmpLogin.UNPassing;
            LbluserId.Text = EmpLogin.IDPassing;
            if (SMSText.Text != "")
            {
                if (SMSText.Text != "")
                {
                    InputLanguage MCL = InputLanguage.CurrentInputLanguage;
                    LblLang.Text = "Language Is : " + MCL.Culture.EnglishName;
                    if (LblLang.Text == "Language Is : English (United States)")
                    {
                        SMSText.MaxLength = 160;
                        Charcount = SMSText.Text.Length;
                        Count.Text = Charcount.ToString() + " " + "/" + " " + "160";
                        int counttext = Convert.ToInt32(Charcount.ToString());
                        if (counttext <= 160)
                        {
                            MN.Text = "1";
                        }
                    }
                    else
                    {
                        SMSText.MaxLength = 70;
                        Charcount = SMSText.Text.Length;
                        Count.Text = Charcount.ToString() + " " + "/" + " " + "70";
                        int counttext = Convert.ToInt32(Charcount.ToString());
                        if (counttext <= 70)
                        {
                            MN.Text = "1";
                        }
                    }
                    BSend.Enabled = true;
                }
                else
                {

                    SMSText.MaxLength = 0;
                    LblLang.Text = "Language";
                    Count.Text = "0";
                    MN.Text = "0";
                    BSend.Enabled = false;
                }
            }
        }
        PortalContext Db = new PortalContext();
        private void Form1_Load(object sender, EventArgs e)
        {
        }
        private void Close_Click(object sender, EventArgs e)
        {
            Close();
        }
        private async void BtnLog_Click(object sender, EventArgs e)
        {
        }
        private void Close_Click_1(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Maximized)
            {
                Hide();
                NI.Visible = true;
            }
        }
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }
        private void SMSText_TextChanged(object sender, EventArgs e)
        {
            if (SMSText.Text != "")
            {
                InputLanguage MCL = InputLanguage.CurrentInputLanguage;
                LblLang.Text = "Language Is : " + MCL.Culture.EnglishName;
                if (LblLang.Text == "Language Is : English (United States)")
                {
                    SMSText.MaxLength = 160;
                    Charcount = SMSText.Text.Length;
                    Count.Text = Charcount.ToString() + " " + "/" + " " + "160";
                    int counttext = Convert.ToInt32(Charcount.ToString());
                    if (counttext <= 160)
                    {
                        MN.Text = "1";
                    }
                }
                else
                {
                    SMSText.MaxLength = 70;
                    Charcount = SMSText.Text.Length;
                    Count.Text = Charcount.ToString() + " " + "/" + " " + "70";
                    int counttext = Convert.ToInt32(Charcount.ToString());
                    if (counttext <= 70)
                    {
                        MN.Text = "1";
                    }
                }
                BSend.Enabled = true;
            }
            else
            {

                SMSText.MaxLength = 0;
                LblLang.Text = "Language";
                Count.Text = "0";
                MN.Text = "0";
                BSend.Enabled = false;
            }
        }
        private async void BSend_Click(object sender, EventArgs e)
        {
            string message = "Are You Shure Want Send Multiple SMS Messages?";
            string title = "Shure ?";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.Yes)
            {
                Microsoft.Office.Interop.Excel.Application XLAPP;
                Microsoft.Office.Interop.Excel.Workbook XLWB;
                Microsoft.Office.Interop.Excel.Worksheet XLWS;
                Microsoft.Office.Interop.Excel.Range XLR;

                int XLRow;

                if (TPath.Text != string.Empty)
                {
                    XLAPP = new Microsoft.Office.Interop.Excel.Application();
                    XLWB = XLAPP.Workbooks.Open(TPath.Text);
                    XLWS = XLWB.Worksheets["Sheet1"];
                    XLR = XLWS.UsedRange;

                    for (XLRow = 1; XLRow <= XLR.Rows.Count; XLRow++)
                    {
                        //DGV.Rows.Add(XLR.Cells[XLRow, 1].Text, XLR.Cells[XLRow, 2].Text);
                        Db = new PortalContext();
                        var Actions = new TransAction();
                        Actions.SMSPhone = "0" + XLR.Cells[XLRow, 1].Text;
                        Actions.SMSBody = SMSText.Text;
                        if (SMSText.RightToLeft == RightToLeft.Yes)
                        {
                            Actions.Languge = "Language Is : Arabic (Egypt)";
                        }
                        else
                        {
                            Actions.Languge = "Language Is : English (United States)";
                        }
                        Actions.UserId = long.Parse(LbluserId.Text.ToString());
                        var T = DateTime.Now.ToLongTimeString();
                        Actions.Date = DateTime.Now;
                        Actions.Time = T;
                        Db.TransActions.Add(Actions);
                        await Db.SaveChangesAsync();
                    }
                    XLWB.Close();
                    XLAPP.Workbooks.Close();
                }
                TPath.Clear();
                SMSText.Clear();
                MessageBox.Show("Your Bulk SMS Uploaded Successfully ...", "Successfully");
            }
            else
            {
                return;
            }
        }
        private void CBUsers_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
        private void Header_Paint(object sender, PaintEventArgs e)
        {

        }
        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
        private void LblName_Click(object sender, EventArgs e)
        {

        }
        private void MenueLogin_Paint(object sender, PaintEventArgs e)
        {

        }
        private void MenueLogout_Paint(object sender, PaintEventArgs e)
        {

        }
        private void LblLogin_Click(object sender, EventArgs e)
        {

        }
        private void Phone_TextChanged(object sender, EventArgs e)
        {

        }
        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void UserName_TextChanged(object sender, EventArgs e)
        {

        }
        private void label2_Click(object sender, EventArgs e)
        {

        }
        private void Bsave_Click(object sender, EventArgs e)
        {
            if (SMSText.Text == "")
            {
                MessageBox.Show("The Message Is Empty", "Error");
            }
            else
            {
            }
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
        private void MN_Click(object sender, EventArgs e)
        {

        }
        private void label9_Click(object sender, EventArgs e)
        {

        }
        private void LblLang_Click(object sender, EventArgs e)
        {

        }
        private void Count_Click(object sender, EventArgs e)
        {

        }
        private void Charcter_Click(object sender, EventArgs e)
        {

        }
        private void PBody_Paint(object sender, PaintEventArgs e)
        {
        }
        private void TPhone_TextChanged(object sender, EventArgs e)
        {

        }
        private void label4_Click(object sender, EventArgs e)
        {

        }
        private void label3_Click(object sender, EventArgs e)
        {

        }
        private void SMSText_KeyPress(object sender, KeyPressEventArgs e)
        {
            string Lang = InputLanguage.CurrentInputLanguage.LayoutName;
            if (Lang == "Arabic (101)")
            {
                SMSText.RightToLeft = RightToLeft.Yes;
            }
            else
            {
                SMSText.RightToLeft = RightToLeft.No;
            }
        }
        private void Btn_Click_1(object sender, EventArgs e)
        {

        }
        private void PBody_Click(object sender, EventArgs e)
        {
        }
        private void TPhone_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
        private void History_Click(object sender, EventArgs e)
        {
            UNPassing = LblName.Text;
            IDPassing = LbluserId.Text;
            var EHF = Application.OpenForms["HistoryBulk"];
            if ((EHF as HistoryBulk) != null)
            {
                if (EHF.WindowState == FormWindowState.Minimized)
                {
                    EHF.WindowState = FormWindowState.Normal;
                }
                else
                {
                    EHF.BringToFront();
                    //Form is already open
                }
            }
            else
            {
                HistoryBulk EHFF = new HistoryBulk();
                EHFF.Show();
                // Form is not open
            }
        }
        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
            foreach (var process in Process.GetProcessesByName("Bulk"))
            {
                process.Kill();
            }
        }
        private void maxmizesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Maximized;
            NI.Visible = false;
        }
        private void NI_Click(object sender, EventArgs e)
        {

        }
        private void NI_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Show();
                this.WindowState = FormWindowState.Maximized;
                NI.Visible = false;
            }
            else
            {
                CMS.Show();
            }
        }
        private void BImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog OFD = new OpenFileDialog();
            OFD.Filter = "ALL Files|*.*|Excel Files |*.xls;*.xlsx;*.xlsm";
            OFD.ShowDialog();
            TPath.Text = OFD.FileName;
        }
    }
}
