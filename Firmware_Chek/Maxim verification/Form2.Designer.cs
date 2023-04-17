namespace Maxim_verification
{
    partial class Form2
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
            this.tbuser = new System.Windows.Forms.TextBox();
            this.tblogin = new System.Windows.Forms.Button();
            this.tbclose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tbuser
            // 
            this.tbuser.Location = new System.Drawing.Point(60, 56);
            this.tbuser.Name = "tbuser";
            this.tbuser.Size = new System.Drawing.Size(100, 20);
            this.tbuser.TabIndex = 0;
            // 
            // tblogin
            // 
            this.tblogin.Location = new System.Drawing.Point(60, 134);
            this.tblogin.Name = "tblogin";
            this.tblogin.Size = new System.Drawing.Size(75, 23);
            this.tblogin.TabIndex = 1;
            this.tblogin.Text = "Login";
            this.tblogin.UseVisualStyleBackColor = true;
            this.tblogin.Click += new System.EventHandler(this.tblogin_Click);
            // 
            // tbclose
            // 
            this.tbclose.Location = new System.Drawing.Point(202, 134);
            this.tbclose.Name = "tbclose";
            this.tbclose.Size = new System.Drawing.Size(75, 23);
            this.tbclose.TabIndex = 2;
            this.tbclose.Text = "Exit";
            this.tbclose.UseVisualStyleBackColor = true;
            this.tbclose.Click += new System.EventHandler(this.tbclose_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDark;
            this.ClientSize = new System.Drawing.Size(392, 234);
            this.ControlBox = false;
            this.Controls.Add(this.tbclose);
            this.Controls.Add(this.tblogin);
            this.Controls.Add(this.tbuser);
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Login Check Rev ID";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbuser;
        private System.Windows.Forms.Button tblogin;
        private System.Windows.Forms.Button tbclose;
    }
}