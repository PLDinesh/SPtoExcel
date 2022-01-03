namespace SPtoExcel
{
    partial class frmMain
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
            this.button1 = new System.Windows.Forms.Button();
            this.txtServerName = new System.Windows.Forms.TextBox();
            this.Server = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.rbnWindowsAuthentication = new System.Windows.Forms.RadioButton();
            this.rbnSqlAuthentication = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtSQLPassword = new System.Windows.Forms.TextBox();
            this.txtSQLUserID = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDatabaseName = new System.Windows.Forms.ComboBox();
            this.grdResultsGrid = new System.Windows.Forms.DataGridView();
            this.btnConnect = new System.Windows.Forms.Button();
            this.txtSQLQuery = new System.Windows.Forms.TextBox();
            this.btnSaveExcel = new System.Windows.Forms.Button();
            this.dlgSaveDialog = new System.Windows.Forms.SaveFileDialog();
            this.label4 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdResultsGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(712, 415);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Execute";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtServerName
            // 
            this.txtServerName.Location = new System.Drawing.Point(223, 35);
            this.txtServerName.Name = "txtServerName";
            this.txtServerName.Size = new System.Drawing.Size(195, 20);
            this.txtServerName.TabIndex = 1;
            this.txtServerName.Text = "dw-devsql";
            // 
            // Server
            // 
            this.Server.AutoSize = true;
            this.Server.Location = new System.Drawing.Point(107, 38);
            this.Server.Name = "Server";
            this.Server.Size = new System.Drawing.Size(38, 13);
            this.Server.TabIndex = 2;
            this.Server.Text = "Server";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(107, 74);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(22, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "DB";
            // 
            // rbnWindowsAuthentication
            // 
            this.rbnWindowsAuthentication.AutoSize = true;
            this.rbnWindowsAuthentication.Checked = true;
            this.rbnWindowsAuthentication.Location = new System.Drawing.Point(24, 31);
            this.rbnWindowsAuthentication.Name = "rbnWindowsAuthentication";
            this.rbnWindowsAuthentication.Size = new System.Drawing.Size(140, 17);
            this.rbnWindowsAuthentication.TabIndex = 3;
            this.rbnWindowsAuthentication.TabStop = true;
            this.rbnWindowsAuthentication.Text = "Windows Authentication";
            this.rbnWindowsAuthentication.UseVisualStyleBackColor = true;
            this.rbnWindowsAuthentication.CheckedChanged += new System.EventHandler(this.rbnWindowsAuthentication_CheckedChanged);
            // 
            // rbnSqlAuthentication
            // 
            this.rbnSqlAuthentication.AutoSize = true;
            this.rbnSqlAuthentication.Location = new System.Drawing.Point(24, 54);
            this.rbnSqlAuthentication.Name = "rbnSqlAuthentication";
            this.rbnSqlAuthentication.Size = new System.Drawing.Size(117, 17);
            this.rbnSqlAuthentication.TabIndex = 3;
            this.rbnSqlAuthentication.Text = "SQL Authentication";
            this.rbnSqlAuthentication.UseVisualStyleBackColor = true;
            this.rbnSqlAuthentication.CheckedChanged += new System.EventHandler(this.rbnSqlAuthentication_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtSQLPassword);
            this.groupBox1.Controls.Add(this.txtSQLUserID);
            this.groupBox1.Controls.Add(this.rbnWindowsAuthentication);
            this.groupBox1.Controls.Add(this.rbnSqlAuthentication);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(110, 106);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(318, 167);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Auth";
            // 
            // txtSQLPassword
            // 
            this.txtSQLPassword.Location = new System.Drawing.Point(104, 104);
            this.txtSQLPassword.Name = "txtSQLPassword";
            this.txtSQLPassword.PasswordChar = '*';
            this.txtSQLPassword.Size = new System.Drawing.Size(100, 20);
            this.txtSQLPassword.TabIndex = 4;
            // 
            // txtSQLUserID
            // 
            this.txtSQLUserID.Location = new System.Drawing.Point(104, 78);
            this.txtSQLUserID.Name = "txtSQLUserID";
            this.txtSQLUserID.Size = new System.Drawing.Size(100, 20);
            this.txtSQLUserID.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(42, 104);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Password";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(52, 81);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "User ID";
            // 
            // txtDatabaseName
            // 
            this.txtDatabaseName.FormattingEnabled = true;
            this.txtDatabaseName.Location = new System.Drawing.Point(223, 65);
            this.txtDatabaseName.Name = "txtDatabaseName";
            this.txtDatabaseName.Size = new System.Drawing.Size(195, 21);
            this.txtDatabaseName.TabIndex = 5;
            this.txtDatabaseName.Text = "CosmosPolaris";
            // 
            // grdResultsGrid
            // 
            this.grdResultsGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdResultsGrid.Location = new System.Drawing.Point(32, 497);
            this.grdResultsGrid.Name = "grdResultsGrid";
            this.grdResultsGrid.Size = new System.Drawing.Size(660, 221);
            this.grdResultsGrid.TabIndex = 6;
            // 
            // btnConnect
            // 
            this.btnConnect.Location = new System.Drawing.Point(476, 137);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(75, 23);
            this.btnConnect.TabIndex = 7;
            this.btnConnect.Text = "Connect";
            this.btnConnect.UseVisualStyleBackColor = true;
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // txtSQLQuery
            // 
            this.txtSQLQuery.Location = new System.Drawing.Point(58, 349);
            this.txtSQLQuery.Multiline = true;
            this.txtSQLQuery.Name = "txtSQLQuery";
            this.txtSQLQuery.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtSQLQuery.Size = new System.Drawing.Size(634, 123);
            this.txtSQLQuery.TabIndex = 8;
            // 
            // btnSaveExcel
            // 
            this.btnSaveExcel.Location = new System.Drawing.Point(712, 520);
            this.btnSaveExcel.Name = "btnSaveExcel";
            this.btnSaveExcel.Size = new System.Drawing.Size(75, 61);
            this.btnSaveExcel.TabIndex = 9;
            this.btnSaveExcel.Text = "Export First ResultSet to Excel";
            this.btnSaveExcel.UseVisualStyleBackColor = true;
            this.btnSaveExcel.Click += new System.EventHandler(this.btnSaveExcel_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(146, 317);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "SQL Query";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(712, 630);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 61);
            this.button2.TabIndex = 9;
            this.button2.Text = "Export All ResultSets to Excel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(611, 196);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(152, 123);
            this.textBox1.TabIndex = 8;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(608, 139);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(185, 39);
            this.label5.TabIndex = 2;
            this.label5.Text = "Work Sheet names (Optional)\r\nPlease enter the unique sheet names \r\nseparated by n" +
                "ew line";
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(799, 730);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btnSaveExcel);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.txtSQLQuery);
            this.Controls.Add(this.btnConnect);
            this.Controls.Add(this.grdResultsGrid);
            this.Controls.Add(this.txtDatabaseName);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Server);
            this.Controls.Add(this.txtServerName);
            this.Controls.Add(this.button1);
            this.Name = "frmMain";
            this.Text = "Export Query/SP output to Excel";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmMain_FormClosed);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdResultsGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtServerName;
        private System.Windows.Forms.Label Server;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton rbnWindowsAuthentication;
        private System.Windows.Forms.RadioButton rbnSqlAuthentication;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox txtDatabaseName;
        private System.Windows.Forms.DataGridView grdResultsGrid;
        private System.Windows.Forms.Button btnConnect;
        private System.Windows.Forms.TextBox txtSQLQuery;
        private System.Windows.Forms.Button btnSaveExcel;
        private System.Windows.Forms.SaveFileDialog dlgSaveDialog;
        private System.Windows.Forms.TextBox txtSQLUserID;
        private System.Windows.Forms.TextBox txtSQLPassword;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label5;
    }
}

