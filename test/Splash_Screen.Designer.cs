namespace helper
{
    partial class Splash_Screen
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
            this.RadEgypt = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.RadUAE = new System.Windows.Forms.RadioButton();
            this.btnChooseLang = new System.Windows.Forms.Button();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.FBDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.OFDialog = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // RadEgypt
            // 
            this.RadEgypt.AutoSize = true;
            this.RadEgypt.Checked = true;
            this.RadEgypt.Location = new System.Drawing.Point(99, 46);
            this.RadEgypt.Name = "RadEgypt";
            this.RadEgypt.Size = new System.Drawing.Size(52, 17);
            this.RadEgypt.TabIndex = 2;
            this.RadEgypt.TabStop = true;
            this.RadEgypt.Text = "Egypt";
            this.RadEgypt.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(96, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(140, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Choose your current country";
            // 
            // RadUAE
            // 
            this.RadUAE.AutoSize = true;
            this.RadUAE.Location = new System.Drawing.Point(163, 46);
            this.RadUAE.Name = "RadUAE";
            this.RadUAE.Size = new System.Drawing.Size(73, 17);
            this.RadUAE.TabIndex = 4;
            this.RadUAE.Text = "UAE/KSA";
            this.RadUAE.UseVisualStyleBackColor = true;
            // 
            // btnChooseLang
            // 
            this.btnChooseLang.Location = new System.Drawing.Point(245, 111);
            this.btnChooseLang.Name = "btnChooseLang";
            this.btnChooseLang.Size = new System.Drawing.Size(75, 23);
            this.btnChooseLang.TabIndex = 5;
            this.btnChooseLang.Text = "Select";
            this.btnChooseLang.UseVisualStyleBackColor = true;
            this.btnChooseLang.Click += new System.EventHandler(this.btnChooseLang_Click);
            // 
            // txtPath
            // 
            this.txtPath.Location = new System.Drawing.Point(23, 74);
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(269, 20);
            this.txtPath.TabIndex = 6;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(298, 74);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(22, 20);
            this.btnBrowse.TabIndex = 7;
            this.btnBrowse.Text = "...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // OFDialog
            // 
            this.OFDialog.FileName = "openFileDialog1";
            this.OFDialog.Filter = "Database File|*.db";
            this.OFDialog.Title = "Choose database file";
            // 
            // Splash_Screen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(329, 149);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.btnChooseLang);
            this.Controls.Add(this.RadUAE);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.RadEgypt);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.Name = "Splash_Screen";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Choose Country";
            this.Activated += new System.EventHandler(this.Splash_Screen_Activated);
            this.Shown += new System.EventHandler(this.Splash_Screen_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RadioButton RadEgypt;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton RadUAE;
        private System.Windows.Forms.Button btnChooseLang;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.FolderBrowserDialog FBDialog;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.OpenFileDialog OFDialog;
    }
}