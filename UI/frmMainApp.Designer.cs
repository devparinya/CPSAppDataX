namespace CPSAppData
{
    partial class frmMainApp
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMainApp));
            btn_setting = new Button();
            btn_report = new Button();
            btn_user_report = new Button();
            brn_advuser = new Button();
            SuspendLayout();
            // 
            // btn_setting
            // 
            btn_setting.Font = new Font("Leelawadee UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            btn_setting.Image = (Image)resources.GetObject("btn_setting.Image");
            btn_setting.ImageAlign = ContentAlignment.TopCenter;
            btn_setting.Location = new Point(361, 278);
            btn_setting.Name = "btn_setting";
            btn_setting.Size = new Size(72, 64);
            btn_setting.TabIndex = 0;
            btn_setting.Text = "Import Data";
            btn_setting.TextAlign = ContentAlignment.BottomCenter;
            btn_setting.UseVisualStyleBackColor = true;
            btn_setting.Click += btn_setting_Click;
            // 
            // btn_report
            // 
            btn_report.Font = new Font("Leelawadee UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            btn_report.Image = (Image)resources.GetObject("btn_report.Image");
            btn_report.ImageAlign = ContentAlignment.TopCenter;
            btn_report.Location = new Point(457, 278);
            btn_report.Name = "btn_report";
            btn_report.Size = new Size(72, 64);
            btn_report.TabIndex = 1;
            btn_report.Text = "Report (Admin)";
            btn_report.TextAlign = ContentAlignment.BottomCenter;
            btn_report.UseVisualStyleBackColor = true;
            btn_report.Click += btn_report_Click;
            // 
            // btn_user_report
            // 
            btn_user_report.Font = new Font("Leelawadee UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            btn_user_report.Image = (Image)resources.GetObject("btn_user_report.Image");
            btn_user_report.ImageAlign = ContentAlignment.TopCenter;
            btn_user_report.Location = new Point(553, 278);
            btn_user_report.Name = "btn_user_report";
            btn_user_report.Size = new Size(72, 64);
            btn_user_report.TabIndex = 1;
            btn_user_report.Text = "Report (User)";
            btn_user_report.TextAlign = ContentAlignment.BottomCenter;
            btn_user_report.UseVisualStyleBackColor = true;
            btn_user_report.Click += btn_user_report_Click;
            // 
            // brn_advuser
            // 
            brn_advuser.Font = new Font("Leelawadee UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            brn_advuser.Image = (Image)resources.GetObject("brn_advuser.Image");
            brn_advuser.ImageAlign = ContentAlignment.TopCenter;
            brn_advuser.Location = new Point(642, 278);
            brn_advuser.Name = "brn_advuser";
            brn_advuser.Size = new Size(90, 64);
            brn_advuser.TabIndex = 1;
            brn_advuser.Text = "Report (Advanced)";
            brn_advuser.TextAlign = ContentAlignment.BottomCenter;
            brn_advuser.UseVisualStyleBackColor = true;
            brn_advuser.Visible = false;
            brn_advuser.Click += brn_advuser_Click;
            // 
            // frmMainApp
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1008, 681);
            Controls.Add(brn_advuser);
            Controls.Add(btn_user_report);
            Controls.Add(btn_report);
            Controls.Add(btn_setting);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "frmMainApp";
            ResumeLayout(false);
        }

        #endregion

        private Button btn_setting;
        private Button btn_report;
        private Button btn_user_report;
        private Button brn_advuser;
    }
}
