using Controler;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Portal.View.Users
{
    public partial class HistoryBulk : Form
    {
        int mov;
        int movx;
        int movy;
        PortalContext Db;
        public HistoryBulk()
        {
            InitializeComponent();
            LblName.Text = Employees.UNPassing;
            LbluserId.Text = Employees.IDPassing;
        }
        private void History_Click(object sender, EventArgs e)
        {
        }
        private void Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }
        private void Header_MouseUp(object sender, MouseEventArgs e)
        {
            mov = 0;
        }
        private void Header_MouseMove(object sender, MouseEventArgs e)
        {
            if (mov == 1)
            {
                this.SetDesktopLocation(MousePosition.X - movx, MousePosition.Y - movy);
            }
        }
        private void Header_MouseDown(object sender, MouseEventArgs e)
        {
            mov = 1;
            movx = e.X;
            movy = e.Y;
        }
        private async void EmpHistories_Load(object sender, EventArgs e)
        {
            DTF.Value = DateTime.Today;
            DTT.Value = DateTime.Today;
            try
            {
                GHLT();
            }
            catch (Exception)
            {
                MessageBox.Show("Server Is Dowen ...","Error");
            }
        }

        //GHLT : Get History List For Today
        public void GHLT()
        {
            Db = new PortalContext();
            int id = Convert.ToInt32(LbluserId.Text);
            var PDb = from hist in Db.Histories
                      where hist.Date == DateTime.Today
                      join USING in Db.Users
                      on hist.UserId equals USING.Id
                      where USING.Id == id
                      select new
                      {
                          hist.SMSPhone,
                          hist.SMSBody,
                          hist.Languge,
                          hist.Status,
                          hist.Result,
                          hist.Date,
                          hist.Time
                      };
            DGV.DataSource = PDb.ToList();
        }

        //GHLT : Get History List For Date Choose
        public void GHLD()
        {
            Db = new PortalContext();
            int id = Convert.ToInt32(LbluserId.Text);
            var PDb = from hist in Db.Histories
                      where hist.Date >= DTF.Value && hist.Date <= DTT.Value
                      join USING in Db.Users
                      on hist.UserId equals USING.Id
                      where USING.Id == id
                      select new
                      {
                          hist.SMSPhone,
                          hist.SMSBody,
                          hist.Languge,
                          hist.Status,
                          hist.Result,
                          hist.Date,
                          hist.Time
                      };
            DGV.DataSource = PDb.ToList();
        }

        //GHLT : Get History Phone List For Today
        public void GHPT()
        {
            Db = new PortalContext();
            int id = Convert.ToInt32(LbluserId.Text);
            var PDb = from hist in Db.Histories
                      where hist.Date == DateTime.Today &&
                            hist.SMSPhone.Contains(TSerching.Text)
                      join USING in Db.Users
                      on hist.UserId equals USING.Id
                      where USING.Id == id
                      select new
                      {
                          hist.SMSPhone,
                          hist.SMSBody,
                          hist.Languge,
                          hist.Status,
                          hist.Result,
                          hist.Date,
                          hist.Time
                      };
            DGV.DataSource = PDb.ToList();
        }

        //GHLT : Get History Phone List For Date Choose
        public void GHPD()
        {
            Db = new PortalContext();
            int id = Convert.ToInt32(LbluserId.Text);
            var PDb = from hist in Db.Histories
                      where hist.Date >= DTF.Value && hist.Date <= DTT.Value &&
                            hist.SMSPhone.Contains(TSerching.Text)
                      join USING in Db.Users
                      on hist.UserId equals USING.Id
                      where USING.Id == id
                      select new
                      {
                          hist.SMSPhone,
                          hist.SMSBody,
                          hist.Languge,
                          hist.Status,
                          hist.Result,
                          hist.Date,
                          hist.Time
                      };
            DGV.DataSource = PDb.ToList();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (TSerching.Text == "" && DTF.Value == DateTime.Today && DTT.Value == DateTime.Today)
            {
                GHLT();
            }
            if (TSerching.Text == "" && DTF.Value <= DateTime.Today && DTT.Value <= DateTime.Today)
            {
                GHLD();
            }
            if (TSerching.Text != "" && DTF.Value == DateTime.Today && DTT.Value == DateTime.Today)
            {
                GHPT();
            }
            if (TSerching.Text != "" && DTF.Value <= DateTime.Today && DTT.Value <= DateTime.Today)
            {
                GHPD();
            }
        }
        private void DTF_KeyDown(object sender, KeyEventArgs e)
        {
        }
        private void DTT_KeyDown(object sender, KeyEventArgs e)
        {
        }
        private void DTF_ValueChanged(object sender, EventArgs e)
        {
        }
        private void DTT_ValueChanged(object sender, EventArgs e)
        {
        }
        private void TSerching_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Key Pres Number Only
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }
        private void TSerching_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (TSerching.Text == "" && DTF.Value == DateTime.Today && DTT.Value == DateTime.Today)
                {
                    GHLT();
                }
                if (TSerching.Text == "" && DTF.Value <= DateTime.Today && DTT.Value <= DateTime.Today)
                {
                    GHLD();
                }
                if (TSerching.Text != "" && DTF.Value == DateTime.Today && DTT.Value == DateTime.Today)
                {
                    GHPT();
                }
                if (TSerching.Text != "" && DTF.Value <= DateTime.Today && DTT.Value <= DateTime.Today)
                {
                    GHPD();
                }
            }
        }
        private void Export_Click(object sender, EventArgs e)
        {

            // creating Excel Application  
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            //// see the excel sheet behind the program  
            //app.Visible = true;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = "Exported From Bulk";
            // storing header part in Excel  
            for (int i = 1; i < DGV.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = DGV.Columns[i - 1].HeaderText;
            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < DGV.Rows.Count - 1; i++)
            {
                for (int j = 0; j < DGV.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = DGV.Rows[i].Cells[j].Value.ToString();
                }
            }
            var ComputerName = Environment.MachineName;
            SaveFileDialog SFD = new SaveFileDialog();
            SFD.InitialDirectory = @"C:\" + ComputerName;
            SFD.RestoreDirectory = true;
            SFD.FileName = ".xlsx";
            SFD.DefaultExt = "xlsx";
            SFD.Filter = "Excel Files (*.xlsx)|*.xlsx";
            string FP = "";
            if (SFD.ShowDialog() == DialogResult.OK)
            {
                FP = SFD.FileName;
            }
            var x = FP;
            // save the application  
            workbook.SaveAs(x, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue
            workbook.Close(true, Type.Missing, Type.Missing);

            // Exit from the application  
            app.Quit();
        }
    }
}
