namespace helper
{
    partial class TranslationUsers
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtTransUserName = new System.Windows.Forms.TextBox();
            this.txtTransPassword = new System.Windows.Forms.MaskedTextBox();
            this.btnLoginTransUser = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Username:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 78);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Password:";
            // 
            // txtTransUserName
            // 
            this.txtTransUserName.Location = new System.Drawing.Point(84, 48);
            this.txtTransUserName.Name = "txtTransUserName";
            this.txtTransUserName.Size = new System.Drawing.Size(205, 20);
            this.txtTransUserName.TabIndex = 0;
            // 
            // txtTransPassword
            // 
            this.txtTransPassword.Location = new System.Drawing.Point(84, 75);
            this.txtTransPassword.Name = "txtTransPassword";
            this.txtTransPassword.PasswordChar = '*';
            this.txtTransPassword.Size = new System.Drawing.Size(205, 20);
            this.txtTransPassword.TabIndex = 3;
            // 
            // btnLoginTransUser
            // 
            this.btnLoginTransUser.Location = new System.Drawing.Point(214, 111);
            this.btnLoginTransUser.Name = "btnLoginTransUser";
            this.btnLoginTransUser.Size = new System.Drawing.Size(75, 23);
            this.btnLoginTransUser.TabIndex = 0;
            this.btnLoginTransUser.Text = "LOG IN";
            this.btnLoginTransUser.UseVisualStyleBackColor = true;
            this.btnLoginTransUser.Click += new System.EventHandler(this.btnLoginTransUser_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(280, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Enter your username and password to edit translation data";
            // 
            // TranslationUsers
            // 
            this.AcceptButton = this.btnLoginTransUser;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(338, 161);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnLoginTransUser);
            this.Controls.Add(this.txtTransPassword);
            this.Controls.Add(this.txtTransUserName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "TranslationUsers";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Translation Users";
            this.Load += new System.EventHandler(this.TranslationUsers_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtTransUserName;
        private System.Windows.Forms.MaskedTextBox txtTransPassword;
        private System.Windows.Forms.Button btnLoginTransUser;
        private System.Windows.Forms.Label label3;
    }
}