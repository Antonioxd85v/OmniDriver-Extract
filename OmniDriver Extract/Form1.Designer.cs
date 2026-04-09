namespace OmniDriver_Extract
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            btnExtrair = new Button();
            folderBrowserDialog1 = new FolderBrowserDialog();
            lblStatus = new Label();
            progressBar1 = new ProgressBar();
            cmbLanguage = new ComboBox();
            txtPath = new TextBox();
            btnBrowse = new Button();
            lblPath = new Label();
            lblLanguage = new Label();
            SuspendLayout();
            // 
            // btnExtrair
            // 
            btnExtrair.BackColor = Color.DodgerBlue;
            btnExtrair.FlatStyle = FlatStyle.Flat;
            btnExtrair.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            btnExtrair.ForeColor = Color.White;
            btnExtrair.Location = new Point(30, 115);
            btnExtrair.Name = "btnExtrair";
            btnExtrair.Size = new Size(460, 45);
            btnExtrair.TabIndex = 7;
            btnExtrair.Text = "Extraer y Organizar Drivers";
            btnExtrair.UseVisualStyleBackColor = false;
            btnExtrair.Click += btnExtrair_Click;
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.ForeColor = Color.Yellow;
            lblStatus.Location = new Point(26, 175);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(43, 20);
            lblStatus.TabIndex = 6;
            lblStatus.Text = "Listo.";
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(30, 195);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(460, 20);
            progressBar1.TabIndex = 5;
            // 
            // cmbLanguage
            // 
            cmbLanguage.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbLanguage.Location = new Point(370, 30);
            cmbLanguage.Name = "cmbLanguage";
            cmbLanguage.Size = new Size(120, 28);
            cmbLanguage.TabIndex = 4;
            cmbLanguage.SelectedIndexChanged += cmbLanguage_SelectedIndexChanged;
            // 
            // txtPath
            // 
            txtPath.Location = new Point(30, 75);
            txtPath.Name = "txtPath";
            txtPath.ReadOnly = true;
            txtPath.Size = new Size(350, 27);
            txtPath.TabIndex = 2;
            // 
            // btnBrowse
            // 
            btnBrowse.BackColor = Color.FromArgb(60, 60, 60);
            btnBrowse.FlatStyle = FlatStyle.Flat;
            btnBrowse.ForeColor = Color.White;
            btnBrowse.Location = new Point(390, 74);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new Size(100, 26);
            btnBrowse.TabIndex = 1;
            btnBrowse.Text = "Examinar...";
            btnBrowse.UseVisualStyleBackColor = false;
            btnBrowse.Click += btnBrowse_Click;
            // 
            // lblPath
            // 
            lblPath.AutoSize = true;
            lblPath.ForeColor = Color.White;
            lblPath.Location = new Point(25, 55);
            lblPath.Name = "lblPath";
            lblPath.Size = new Size(168, 20);
            lblPath.TabIndex = 3;
            lblPath.Text = "Ubicación de guardado:";
            // 
            // lblLanguage
            // 
            lblLanguage.AutoSize = true;
            lblLanguage.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
            lblLanguage.ForeColor = Color.DodgerBlue;
            lblLanguage.Location = new Point(370, 10);
            lblLanguage.Name = "lblLanguage";
            lblLanguage.Size = new Size(62, 20);
            lblLanguage.TabIndex = 0;
            lblLanguage.Text = "Idioma:";
            // 
            // Form1
            // 
            BackColor = Color.FromArgb(35, 35, 35);
            ClientSize = new Size(520, 240);
            Controls.Add(lblLanguage);
            Controls.Add(btnBrowse);
            Controls.Add(txtPath);
            Controls.Add(lblPath);
            Controls.Add(cmbLanguage);
            Controls.Add(progressBar1);
            Controls.Add(lblStatus);
            Controls.Add(btnExtrair);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "OmniDriver Extract - Professional Technician";
            ResumeLayout(false);
            PerformLayout();
        }

        private System.Windows.Forms.Button btnExtrair;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ComboBox cmbLanguage;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.Label lblLanguage;
    }
}