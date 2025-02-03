namespace CPSAppData.UI.Setting
{
    partial class frmAdminData
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
            groupBox1 = new GroupBox();
            txt_show = new TextBox();
            btn_clear = new Button();
            btn_decrypt = new Button();
            btn_encrypt = new Button();
            label1 = new Label();
            txt_data = new TextBox();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(txt_show);
            groupBox1.Controls.Add(btn_clear);
            groupBox1.Controls.Add(btn_decrypt);
            groupBox1.Controls.Add(btn_encrypt);
            groupBox1.Controls.Add(label1);
            groupBox1.Controls.Add(txt_data);
            groupBox1.Location = new Point(12, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(595, 177);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            // 
            // txt_show
            // 
            txt_show.BorderStyle = BorderStyle.FixedSingle;
            txt_show.ForeColor = Color.Navy;
            txt_show.Location = new Point(37, 102);
            txt_show.Multiline = true;
            txt_show.Name = "txt_show";
            txt_show.Size = new Size(522, 58);
            txt_show.TabIndex = 3;
            // 
            // btn_clear
            // 
            btn_clear.Location = new Point(37, 71);
            btn_clear.Name = "btn_clear";
            btn_clear.Size = new Size(68, 25);
            btn_clear.TabIndex = 2;
            btn_clear.Text = "ล้างข้อมูล";
            btn_clear.UseVisualStyleBackColor = true;
            btn_clear.Click += btn_clear_Click;
            // 
            // btn_decrypt
            // 
            btn_decrypt.Location = new Point(417, 71);
            btn_decrypt.Name = "btn_decrypt";
            btn_decrypt.Size = new Size(68, 25);
            btn_decrypt.TabIndex = 2;
            btn_decrypt.Text = "ถอดรหัส";
            btn_decrypt.UseVisualStyleBackColor = true;
            btn_decrypt.Click += btn_decrypt_Click;
            // 
            // btn_encrypt
            // 
            btn_encrypt.Location = new Point(491, 71);
            btn_encrypt.Name = "btn_encrypt";
            btn_encrypt.Size = new Size(68, 25);
            btn_encrypt.TabIndex = 2;
            btn_encrypt.Text = "เข้ารหัส";
            btn_encrypt.UseVisualStyleBackColor = true;
            btn_encrypt.Click += btn_encrypt_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(37, 20);
            label1.Name = "label1";
            label1.Size = new Size(37, 17);
            label1.TabIndex = 1;
            label1.Text = "ข้อมูล";
            // 
            // txt_data
            // 
            txt_data.BorderStyle = BorderStyle.FixedSingle;
            txt_data.Location = new Point(37, 40);
            txt_data.Name = "txt_data";
            txt_data.Size = new Size(522, 25);
            txt_data.TabIndex = 0;
            // 
            // frmAdminData
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(634, 206);
            Controls.Add(groupBox1);
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            Name = "frmAdminData";
            ShowIcon = false;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private GroupBox groupBox1;
        private TextBox txt_data;
        private Button btn_encrypt;
        private Label label1;
        private Button btn_decrypt;
        private TextBox txt_show;
        private Button btn_clear;
    }
}