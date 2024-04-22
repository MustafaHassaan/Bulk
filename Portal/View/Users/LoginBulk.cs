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
    public partial class EmpLogin : Form
    {
        int mov;
        int movx;
        int movy;
        PortalContext DbE;
        public static string UNPassing;
        public static string IDPassing;
        public EmpLogin()
        {
            InitializeComponent();
        }

        private void ShowPassword_Click(object sender, EventArgs e)
        {
            if (TPassword.UseSystemPasswordChar == true)
            {
                TPassword.UseSystemPasswordChar = false;
            }
            else
            {
                TPassword.UseSystemPasswordChar = true;
            }
        }
        private void Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void BSend_Click(object sender, EventArgs e)
        {
        }
        private void EmpLogin_Load(object sender, EventArgs e)
        {

        }
        private void BEL_Click(object sender, EventArgs e)
        {
            DbE = new PortalContext();
            if (string.IsNullOrEmpty(TUserName.Text) || string.IsNullOrEmpty(TPassword.Text))
            {
                MessageBox.Show("Please Insert UserName & Phone ...", "Error");
                return;
            }
            else if (!string.IsNullOrEmpty(TUserName.Text) && !string.IsNullOrEmpty(TPassword.Text))
            {
                try
                {
                    foreach (var user in DbE.Users)
                    {
                        var x = user.PositionId;
                        if (user.UserName == TUserName.Text && user.Password == TPassword.Text && user.PositionId != 2)
                        {
                            MessageBox.Show("Login Is Successfullu ...", "Success");
                            UNPassing = user.UserName;
                            IDPassing = Convert.ToString(user.Id);
                            Employees FE = new Employees();
                            FE.Show();
                            this.Hide();
                            return;
                        }
                        else if (user.UserName == TUserName.Text && user.Password == TPassword.Text && user.PositionId == 2)
                        {
                            MessageBox.Show("Dont Have Permotion To Open This Window ...", "Warning");
                            return;
                        }
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Server Is Dowen ...", "Warning");
                    return;
                }
            }
            if(!string.IsNullOrEmpty(TUserName.Text) && !string.IsNullOrEmpty(TPassword.Text))
            {
                MessageBox.Show("Login Is Faild Please Try Again OR Signe UP ...", "Faild");
                return;
            }
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
        private void TUserName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DbE = new PortalContext();
                if (string.IsNullOrEmpty(TUserName.Text) || string.IsNullOrEmpty(TPassword.Text))
                {
                    MessageBox.Show("Please Insert UserName & Phone ...", "Error");
                    return;
                }
                else if (!string.IsNullOrEmpty(TUserName.Text) && !string.IsNullOrEmpty(TPassword.Text))
                {
                    try
                    {
                        foreach (var user in DbE.Users)
                        {
                            var x = user.PositionId;
                            if (user.UserName == TUserName.Text && user.Password == TPassword.Text && user.PositionId != 2)
                            {
                                MessageBox.Show("Login Is Successfullu ...", "Success");
                                UNPassing = user.UserName;
                                IDPassing = Convert.ToString(user.Id);
                                Employees FE = new Employees();
                                FE.Show();
                                this.Hide();
                                return;
                            }
                            else if (user.UserName == TUserName.Text && user.Password == TPassword.Text && user.PositionId == 2)
                            {
                                MessageBox.Show("Dont Have Permotion To Open This Window ...", "Warning");
                                return;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Server Is Dowen ...", "Warning");
                        return;
                    }
                }
                if (!string.IsNullOrEmpty(TUserName.Text) && !string.IsNullOrEmpty(TPassword.Text))
                {
                    MessageBox.Show("Login Is Faild Please Try Again OR Signe UP ...", "Faild");
                    return;
                }
            }
        }
        private void TPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DbE = new PortalContext();
                if (string.IsNullOrEmpty(TUserName.Text) || string.IsNullOrEmpty(TPassword.Text))
                {
                    MessageBox.Show("Please Insert UserName & Phone ...", "Error");
                    return;
                }
                else if (!string.IsNullOrEmpty(TUserName.Text) && !string.IsNullOrEmpty(TPassword.Text))
                {
                    try
                    {
                        foreach (var user in DbE.Users)
                        {
                            var x = user.PositionId;
                            if (user.UserName == TUserName.Text && user.Password == TPassword.Text && user.PositionId != 2)
                            {
                                MessageBox.Show("Login Is Successfullu ...", "Success");
                                UNPassing = user.UserName;
                                IDPassing = Convert.ToString(user.Id);
                                Employees FE = new Employees();
                                FE.Show();
                                this.Hide();
                                return;
                            }
                            else if (user.UserName == TUserName.Text && user.Password == TPassword.Text && user.PositionId == 2)
                            {
                                MessageBox.Show("Dont Have Permotion To Open This Window ...", "Warning");
                                return;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Server Is Dowen ...", "Warning");
                        return;
                    }
                }
                if (!string.IsNullOrEmpty(TUserName.Text) && !string.IsNullOrEmpty(TPassword.Text))
                {
                    MessageBox.Show("Login Is Faild Please Try Again OR Signe UP ...", "Faild");
                    return;
                }
            }
        }
    }
}
