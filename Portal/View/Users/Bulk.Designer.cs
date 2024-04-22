namespace Portal
{
    partial class Employees
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Employees));
            this.Header = new System.Windows.Forms.Panel();
            this.Logo = new System.Windows.Forms.PictureBox();
            this.History = new System.Windows.Forms.PictureBox();
            this.LbluserId = new System.Windows.Forms.Label();
            this.LblName = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.Close = new System.Windows.Forms.PictureBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.LblPath = new System.Windows.Forms.Label();
            this.LblSMS = new System.Windows.Forms.Label();
            this.TPath = new System.Windows.Forms.TextBox();
            this.SMSText = new System.Windows.Forms.RichTextBox();
            this.LblCC = new System.Windows.Forms.Label();
            this.Count = new System.Windows.Forms.Label();
            this.LblLang = new System.Windows.Forms.Label();
            this.LblMS = new System.Windows.Forms.Label();
            this.MN = new System.Windows.Forms.Label();
            this.PBody = new System.Windows.Forms.Panel();
            this.BImport = new System.Windows.Forms.Button();
            this.BSend = new System.Windows.Forms.Button();
            this.CMS = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.maxmizesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.NI = new System.Windows.Forms.NotifyIcon(this.components);
            this.Header.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Logo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.History)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Close)).BeginInit();
            this.PBody.SuspendLayout();
            this.CMS.SuspendLayout();
            this.SuspendLayout();
            // 
            // Header
            // 
            this.Header.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.Header.Controls.Add(this.Logo);
            this.Header.Controls.Add(this.History);
            this.Header.Controls.Add(this.LbluserId);
            this.Header.Controls.Add(this.LblName);
            this.Header.Controls.Add(this.pictureBox2);
            this.Header.Controls.Add(this.Close);
            this.Header.Controls.Add(this.panel1);
            this.Header.Dock = System.Windows.Forms.DockStyle.Top;
            this.Header.Location = new System.Drawing.Point(0, 0);
            this.Header.Name = "Header";
            this.Header.Size = new System.Drawing.Size(799, 50);
            this.Header.TabIndex = 6;
            this.Header.Paint += new System.Windows.Forms.PaintEventHandler(this.Header_Paint);
            // 
            // Logo
            // 
            this.Logo.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.Logo.BackColor = System.Drawing.Color.Transparent;
            this.Logo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Logo.Image = global::Bulk.Properties.Resources.cardboard_box_96px;
            this.Logo.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.Logo.Location = new System.Drawing.Point(12, 7);
            this.Logo.Name = "Logo";
            this.Logo.Size = new System.Drawing.Size(39, 37);
            this.Logo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.Logo.TabIndex = 81;
            this.Logo.TabStop = false;
            // 
            // History
            // 
            this.History.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.History.BackColor = System.Drawing.Color.Transparent;
            this.History.Cursor = System.Windows.Forms.Cursors.Hand;
            this.History.Image = global::Bulk.Properties.Resources.past_96px;
            this.History.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.History.Location = new System.Drawing.Point(57, 7);
            this.History.Name = "History";
            this.History.Size = new System.Drawing.Size(39, 37);
            this.History.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.History.TabIndex = 78;
            this.History.TabStop = false;
            this.History.Click += new System.EventHandler(this.History_Click);
            // 
            // LbluserId
            // 
            this.LbluserId.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.LbluserId.AutoSize = true;
            this.LbluserId.BackColor = System.Drawing.Color.Transparent;
            this.LbluserId.Cursor = System.Windows.Forms.Cursors.Default;
            this.LbluserId.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold);
            this.LbluserId.ForeColor = System.Drawing.Color.MediumOrchid;
            this.LbluserId.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.LbluserId.Location = new System.Drawing.Point(315, 14);
            this.LbluserId.Name = "LbluserId";
            this.LbluserId.Size = new System.Drawing.Size(21, 24);
            this.LbluserId.TabIndex = 77;
            this.LbluserId.Text = "0";
            this.LbluserId.Visible = false;
            // 
            // LblName
            // 
            this.LblName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.LblName.AutoSize = true;
            this.LblName.BackColor = System.Drawing.Color.Transparent;
            this.LblName.Cursor = System.Windows.Forms.Cursors.Default;
            this.LblName.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold);
            this.LblName.ForeColor = System.Drawing.Color.MediumOrchid;
            this.LblName.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.LblName.Location = new System.Drawing.Point(102, 14);
            this.LblName.Name = "LblName";
            this.LblName.Size = new System.Drawing.Size(207, 24);
            this.LblName.TabIndex = 74;
            this.LblName.Text = "SMS Employee Portal";
            this.LblName.Click += new System.EventHandler(this.LblName_Click);
            // 
            // pictureBox2
            // 
            this.pictureBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox2.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.pictureBox2.Location = new System.Drawing.Point(712, 10);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(39, 37);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 74;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.Click += new System.EventHandler(this.pictureBox2_Click);
            // 
            // Close
            // 
            this.Close.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Close.BackColor = System.Drawing.Color.Transparent;
            this.Close.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Close.Image = ((System.Drawing.Image)(resources.GetObject("Close.Image")));
            this.Close.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.Close.Location = new System.Drawing.Point(757, 10);
            this.Close.Name = "Close";
            this.Close.Size = new System.Drawing.Size(39, 37);
            this.Close.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.Close.TabIndex = 72;
            this.Close.TabStop = false;
            this.Close.Click += new System.EventHandler(this.Close_Click_1);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Purple;
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(799, 3);
            this.panel1.TabIndex = 71;
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // LblPath
            // 
            this.LblPath.AutoSize = true;
            this.LblPath.Cursor = System.Windows.Forms.Cursors.Default;
            this.LblPath.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold);
            this.LblPath.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.LblPath.Location = new System.Drawing.Point(65, 99);
            this.LblPath.Name = "LblPath";
            this.LblPath.Size = new System.Drawing.Size(70, 24);
            this.LblPath.TabIndex = 84;
            this.LblPath.Text = "Path : ";
            this.LblPath.Click += new System.EventHandler(this.label3_Click);
            // 
            // LblSMS
            // 
            this.LblSMS.AutoSize = true;
            this.LblSMS.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold);
            this.LblSMS.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.LblSMS.Location = new System.Drawing.Point(65, 146);
            this.LblSMS.Name = "LblSMS";
            this.LblSMS.Size = new System.Drawing.Size(74, 24);
            this.LblSMS.TabIndex = 83;
            this.LblSMS.Text = "Body : ";
            this.LblSMS.Click += new System.EventHandler(this.label4_Click);
            // 
            // TPath
            // 
            this.TPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TPath.Enabled = false;
            this.TPath.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TPath.Location = new System.Drawing.Point(141, 98);
            this.TPath.MaxLength = 500;
            this.TPath.Multiline = true;
            this.TPath.Name = "TPath";
            this.TPath.ReadOnly = true;
            this.TPath.Size = new System.Drawing.Size(421, 29);
            this.TPath.TabIndex = 85;
            this.TPath.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TPhone_KeyPress);
            // 
            // SMSText
            // 
            this.SMSText.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.SMSText.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SMSText.Location = new System.Drawing.Point(141, 140);
            this.SMSText.MaxLength = 1000;
            this.SMSText.Name = "SMSText";
            this.SMSText.Size = new System.Drawing.Size(540, 211);
            this.SMSText.TabIndex = 86;
            this.SMSText.Text = "";
            this.SMSText.TextChanged += new System.EventHandler(this.SMSText_TextChanged);
            this.SMSText.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.SMSText_KeyPress);
            // 
            // LblCC
            // 
            this.LblCC.AutoSize = true;
            this.LblCC.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold);
            this.LblCC.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.LblCC.Location = new System.Drawing.Point(137, 370);
            this.LblCC.Name = "LblCC";
            this.LblCC.Size = new System.Drawing.Size(173, 24);
            this.LblCC.TabIndex = 87;
            this.LblCC.Text = "Character Count :";
            this.LblCC.Click += new System.EventHandler(this.Charcter_Click);
            // 
            // Count
            // 
            this.Count.AutoSize = true;
            this.Count.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold);
            this.Count.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.Count.Location = new System.Drawing.Point(316, 370);
            this.Count.Name = "Count";
            this.Count.Size = new System.Drawing.Size(21, 24);
            this.Count.TabIndex = 88;
            this.Count.Text = "0";
            this.Count.Click += new System.EventHandler(this.Count_Click);
            // 
            // LblLang
            // 
            this.LblLang.AutoSize = true;
            this.LblLang.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold);
            this.LblLang.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.LblLang.Location = new System.Drawing.Point(137, 433);
            this.LblLang.Name = "LblLang";
            this.LblLang.Size = new System.Drawing.Size(100, 24);
            this.LblLang.TabIndex = 89;
            this.LblLang.Text = "Language";
            this.LblLang.Click += new System.EventHandler(this.LblLang_Click);
            // 
            // LblMS
            // 
            this.LblMS.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.LblMS.AutoSize = true;
            this.LblMS.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold);
            this.LblMS.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.LblMS.Location = new System.Drawing.Point(456, 371);
            this.LblMS.Name = "LblMS";
            this.LblMS.Size = new System.Drawing.Size(176, 24);
            this.LblMS.TabIndex = 90;
            this.LblMS.Text = "Message Number :";
            this.LblMS.Click += new System.EventHandler(this.label9_Click);
            // 
            // MN
            // 
            this.MN.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.MN.AutoSize = true;
            this.MN.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold);
            this.MN.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.MN.Location = new System.Drawing.Point(635, 371);
            this.MN.Name = "MN";
            this.MN.Size = new System.Drawing.Size(21, 24);
            this.MN.TabIndex = 91;
            this.MN.Text = "0";
            this.MN.Click += new System.EventHandler(this.MN_Click);
            // 
            // PBody
            // 
            this.PBody.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(188)))), ((int)(((byte)(231)))));
            this.PBody.Controls.Add(this.BImport);
            this.PBody.Controls.Add(this.BSend);
            this.PBody.Controls.Add(this.MN);
            this.PBody.Controls.Add(this.LblMS);
            this.PBody.Controls.Add(this.LblLang);
            this.PBody.Controls.Add(this.Count);
            this.PBody.Controls.Add(this.LblCC);
            this.PBody.Controls.Add(this.SMSText);
            this.PBody.Controls.Add(this.TPath);
            this.PBody.Controls.Add(this.LblSMS);
            this.PBody.Controls.Add(this.LblPath);
            this.PBody.Dock = System.Windows.Forms.DockStyle.Fill;
            this.PBody.Location = new System.Drawing.Point(0, 50);
            this.PBody.Name = "PBody";
            this.PBody.Size = new System.Drawing.Size(799, 625);
            this.PBody.TabIndex = 9;
            this.PBody.Click += new System.EventHandler(this.PBody_Click);
            this.PBody.Paint += new System.Windows.Forms.PaintEventHandler(this.PBody_Paint);
            // 
            // BImport
            // 
            this.BImport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BImport.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BImport.FlatAppearance.BorderSize = 0;
            this.BImport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BImport.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BImport.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.BImport.Image = global::Bulk.Properties.Resources.import_32px;
            this.BImport.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.BImport.Location = new System.Drawing.Point(568, 88);
            this.BImport.Name = "BImport";
            this.BImport.Size = new System.Drawing.Size(113, 46);
            this.BImport.TabIndex = 93;
            this.BImport.Text = "Import";
            this.BImport.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BImport.UseVisualStyleBackColor = true;
            this.BImport.Click += new System.EventHandler(this.BImport_Click);
            // 
            // BSend
            // 
            this.BSend.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.BSend.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BSend.Enabled = false;
            this.BSend.FlatAppearance.BorderSize = 0;
            this.BSend.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BSend.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BSend.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.BSend.Image = ((System.Drawing.Image)(resources.GetObject("BSend.Image")));
            this.BSend.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.BSend.Location = new System.Drawing.Point(340, 567);
            this.BSend.Name = "BSend";
            this.BSend.Size = new System.Drawing.Size(113, 46);
            this.BSend.TabIndex = 92;
            this.BSend.Text = "Send";
            this.BSend.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BSend.UseVisualStyleBackColor = true;
            this.BSend.Click += new System.EventHandler(this.BSend_Click);
            // 
            // CMS
            // 
            this.CMS.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.CMS.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CMS.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.maxmizesToolStripMenuItem,
            this.closeToolStripMenuItem});
            this.CMS.Name = "CMS";
            this.CMS.Size = new System.Drawing.Size(160, 52);
            // 
            // maxmizesToolStripMenuItem
            // 
            this.maxmizesToolStripMenuItem.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.maxmizesToolStripMenuItem.Image = global::Bulk.Properties.Resources.maximize_window_96px;
            this.maxmizesToolStripMenuItem.Name = "maxmizesToolStripMenuItem";
            this.maxmizesToolStripMenuItem.Size = new System.Drawing.Size(159, 24);
            this.maxmizesToolStripMenuItem.Text = "Maxmizes";
            this.maxmizesToolStripMenuItem.Click += new System.EventHandler(this.maxmizesToolStripMenuItem_Click);
            // 
            // closeToolStripMenuItem
            // 
            this.closeToolStripMenuItem.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.closeToolStripMenuItem.Image = global::Bulk.Properties.Resources.cancel_96px;
            this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
            this.closeToolStripMenuItem.Size = new System.Drawing.Size(159, 24);
            this.closeToolStripMenuItem.Text = "Close";
            this.closeToolStripMenuItem.Click += new System.EventHandler(this.closeToolStripMenuItem_Click);
            // 
            // NI
            // 
            this.NI.ContextMenuStrip = this.CMS;
            this.NI.Icon = ((System.Drawing.Icon)(resources.GetObject("NI.Icon")));
            this.NI.Text = "Portal";
            this.NI.Click += new System.EventHandler(this.NI_Click);
            this.NI.MouseClick += new System.Windows.Forms.MouseEventHandler(this.NI_MouseClick);
            // 
            // Employees
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(188)))), ((int)(((byte)(231)))));
            this.ClientSize = new System.Drawing.Size(799, 675);
            this.Controls.Add(this.PBody);
            this.Controls.Add(this.Header);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Employees";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Employees";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Header.ResumeLayout(false);
            this.Header.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Logo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.History)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Close)).EndInit();
            this.PBody.ResumeLayout(false);
            this.PBody.PerformLayout();
            this.CMS.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel Header;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox Close;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label LblName;
        private System.Windows.Forms.Label LbluserId;
        private System.Windows.Forms.Label LblPath;
        private System.Windows.Forms.Label LblSMS;
        private System.Windows.Forms.TextBox TPath;
        private System.Windows.Forms.RichTextBox SMSText;
        private System.Windows.Forms.Label LblCC;
        private System.Windows.Forms.Label Count;
        private System.Windows.Forms.Label LblLang;
        private System.Windows.Forms.Label LblMS;
        private System.Windows.Forms.Label MN;
        private System.Windows.Forms.Button BSend;
        private System.Windows.Forms.Panel PBody;
        private System.Windows.Forms.PictureBox History;
        private System.Windows.Forms.ContextMenuStrip CMS;
        private System.Windows.Forms.ToolStripMenuItem maxmizesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem closeToolStripMenuItem;
        private System.Windows.Forms.NotifyIcon NI;
        private System.Windows.Forms.PictureBox Logo;
        private System.Windows.Forms.Button BImport;
    }
}

