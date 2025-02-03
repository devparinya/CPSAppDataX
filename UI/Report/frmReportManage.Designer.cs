namespace CPSAppData.UI.Report
{
    partial class frmReportManage
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmReportManage));
            tabControl1 = new TabControl();
            tabPage2 = new TabPage();
            groupBox6 = new GroupBox();
            chk_merge = new CheckBox();
            btn_C2TableRangePDF = new Button();
            btn_TableRangePDF = new Button();
            btn_export_Excel = new Button();
            btn_C2RangePDF = new Button();
            dataGridShow = new DataGridView();
            groupBox5 = new GroupBox();
            brn_getwithrange = new Button();
            label28 = new Label();
            label26 = new Label();
            label27 = new Label();
            label25 = new Label();
            txt_endlednumber = new TextBox();
            txt_endworkno = new TextBox();
            txt_startlednumber = new TextBox();
            txt_startworkno = new TextBox();
            tabPage3 = new TabPage();
            groupBox4 = new GroupBox();
            txt_maxmonth = new TextBox();
            label30 = new Label();
            label24 = new Label();
            txt_firstInstalldate = new MaskedTextBox();
            label23 = new Label();
            txt_dateatcalculate = new MaskedTextBox();
            txt_datefest = new MaskedTextBox();
            label31 = new Label();
            label29 = new Label();
            btn_save = new Button();
            label22 = new Label();
            txt_festName = new TextBox();
            label8 = new Label();
            label7 = new Label();
            txt_festNo = new TextBox();
            label6 = new Label();
            groupBox1 = new GroupBox();
            btn_table_cal = new Button();
            btn_path_c2 = new Button();
            label2 = new Label();
            label1 = new Label();
            txt_path_table = new TextBox();
            txt_path_c2 = new TextBox();
            tabControl1.SuspendLayout();
            tabPage2.SuspendLayout();
            groupBox6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridShow).BeginInit();
            groupBox5.SuspendLayout();
            tabPage3.SuspendLayout();
            groupBox4.SuspendLayout();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Controls.Add(tabPage3);
            tabControl1.Dock = DockStyle.Fill;
            tabControl1.Location = new Point(0, 0);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(1008, 681);
            tabControl1.TabIndex = 2;
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(groupBox6);
            tabPage2.Controls.Add(groupBox5);
            tabPage2.Location = new Point(4, 26);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(1000, 651);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "Print With Range Data";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox6
            // 
            groupBox6.Controls.Add(chk_merge);
            groupBox6.Controls.Add(btn_C2TableRangePDF);
            groupBox6.Controls.Add(btn_TableRangePDF);
            groupBox6.Controls.Add(btn_export_Excel);
            groupBox6.Controls.Add(btn_C2RangePDF);
            groupBox6.Controls.Add(dataGridShow);
            groupBox6.Location = new Point(20, 117);
            groupBox6.Name = "groupBox6";
            groupBox6.Size = new Size(956, 500);
            groupBox6.TabIndex = 1;
            groupBox6.TabStop = false;
            // 
            // chk_merge
            // 
            chk_merge.AutoSize = true;
            chk_merge.Location = new Point(434, 468);
            chk_merge.Name = "chk_merge";
            chk_merge.Size = new Size(142, 21);
            chk_merge.TabIndex = 5;
            chk_merge.Text = "รวมเอกสารเป็นไฟล์เดียว";
            chk_merge.UseVisualStyleBackColor = true;
            // 
            // btn_C2TableRangePDF
            // 
            btn_C2TableRangePDF.Image = (Image)resources.GetObject("btn_C2TableRangePDF.Image");
            btn_C2TableRangePDF.Location = new Point(807, 460);
            btn_C2TableRangePDF.Name = "btn_C2TableRangePDF";
            btn_C2TableRangePDF.Size = new Size(132, 34);
            btn_C2TableRangePDF.TabIndex = 4;
            btn_C2TableRangePDF.Text = "ตารางเจรจา&&C2";
            btn_C2TableRangePDF.TextAlign = ContentAlignment.MiddleLeft;
            btn_C2TableRangePDF.TextImageRelation = TextImageRelation.ImageBeforeText;
            btn_C2TableRangePDF.UseVisualStyleBackColor = true;
            btn_C2TableRangePDF.Click += btn_C2TableRangePDF_Click;
            // 
            // btn_TableRangePDF
            // 
            btn_TableRangePDF.Image = (Image)resources.GetObject("btn_TableRangePDF.Image");
            btn_TableRangePDF.Location = new Point(699, 460);
            btn_TableRangePDF.Name = "btn_TableRangePDF";
            btn_TableRangePDF.Size = new Size(102, 34);
            btn_TableRangePDF.TabIndex = 3;
            btn_TableRangePDF.Text = "ตารางเจรจา";
            btn_TableRangePDF.TextAlign = ContentAlignment.MiddleLeft;
            btn_TableRangePDF.TextImageRelation = TextImageRelation.ImageBeforeText;
            btn_TableRangePDF.UseVisualStyleBackColor = true;
            btn_TableRangePDF.Click += btn_TableRangePDF_Click;
            // 
            // btn_export_Excel
            // 
            btn_export_Excel.Image = (Image)resources.GetObject("btn_export_Excel.Image");
            btn_export_Excel.Location = new Point(21, 460);
            btn_export_Excel.Name = "btn_export_Excel";
            btn_export_Excel.Size = new Size(115, 29);
            btn_export_Excel.TabIndex = 2;
            btn_export_Excel.Text = "Export Excel";
            btn_export_Excel.TextAlign = ContentAlignment.MiddleLeft;
            btn_export_Excel.TextImageRelation = TextImageRelation.ImageBeforeText;
            btn_export_Excel.UseVisualStyleBackColor = true;
            btn_export_Excel.Click += btn_export_Excel_Click;
            // 
            // btn_C2RangePDF
            // 
            btn_C2RangePDF.Image = (Image)resources.GetObject("btn_C2RangePDF.Image");
            btn_C2RangePDF.Location = new Point(592, 460);
            btn_C2RangePDF.Name = "btn_C2RangePDF";
            btn_C2RangePDF.Size = new Size(101, 34);
            btn_C2RangePDF.TabIndex = 2;
            btn_C2RangePDF.Text = "เอกสาร C2";
            btn_C2RangePDF.TextAlign = ContentAlignment.MiddleLeft;
            btn_C2RangePDF.TextImageRelation = TextImageRelation.ImageBeforeText;
            btn_C2RangePDF.UseVisualStyleBackColor = true;
            btn_C2RangePDF.Click += btn_C2RangePDF_Click;
            // 
            // dataGridShow
            // 
            dataGridShow.AllowUserToAddRows = false;
            dataGridShow.AllowUserToDeleteRows = false;
            dataGridShow.AllowUserToResizeRows = false;
            dataGridShow.Anchor = AnchorStyles.None;
            dataGridShow.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridShow.Location = new Point(21, 24);
            dataGridShow.MultiSelect = false;
            dataGridShow.Name = "dataGridShow";
            dataGridShow.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dataGridShow.RowHeadersVisible = false;
            dataGridShow.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridShow.Size = new Size(918, 430);
            dataGridShow.TabIndex = 0;
            // 
            // groupBox5
            // 
            groupBox5.Controls.Add(brn_getwithrange);
            groupBox5.Controls.Add(label28);
            groupBox5.Controls.Add(label26);
            groupBox5.Controls.Add(label27);
            groupBox5.Controls.Add(label25);
            groupBox5.Controls.Add(txt_endlednumber);
            groupBox5.Controls.Add(txt_endworkno);
            groupBox5.Controls.Add(txt_startlednumber);
            groupBox5.Controls.Add(txt_startworkno);
            groupBox5.Location = new Point(20, 6);
            groupBox5.Name = "groupBox5";
            groupBox5.Size = new Size(956, 105);
            groupBox5.TabIndex = 0;
            groupBox5.TabStop = false;
            // 
            // brn_getwithrange
            // 
            brn_getwithrange.Location = new Point(452, 61);
            brn_getwithrange.Name = "brn_getwithrange";
            brn_getwithrange.Size = new Size(75, 27);
            brn_getwithrange.TabIndex = 5;
            brn_getwithrange.Text = "แสดงข้อมูล";
            brn_getwithrange.UseVisualStyleBackColor = true;
            brn_getwithrange.Click += brn_getwithrange_Click;
            // 
            // label28
            // 
            label28.AutoSize = true;
            label28.Location = new Point(251, 65);
            label28.Name = "label28";
            label28.Size = new Size(20, 17);
            label28.TabIndex = 1;
            label28.Text = "ถึง";
            // 
            // label26
            // 
            label26.AutoSize = true;
            label26.Location = new Point(251, 29);
            label26.Name = "label26";
            label26.Size = new Size(20, 17);
            label26.TabIndex = 1;
            label26.Text = "ถึง";
            // 
            // label27
            // 
            label27.AutoSize = true;
            label27.Location = new Point(9, 65);
            label27.Name = "label27";
            label27.Size = new Size(75, 17);
            label27.TabIndex = 1;
            label27.Text = "จากลำดับกรม";
            // 
            // label25
            // 
            label25.AutoSize = true;
            label25.Location = new Point(21, 32);
            label25.Name = "label25";
            label25.Size = new Size(63, 17);
            label25.TabIndex = 1;
            label25.Text = "จากลำดับที่";
            // 
            // txt_endlednumber
            // 
            txt_endlednumber.BorderStyle = BorderStyle.FixedSingle;
            txt_endlednumber.CharacterCasing = CharacterCasing.Upper;
            txt_endlednumber.Font = new Font("Leelawadee UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            txt_endlednumber.Location = new Point(278, 61);
            txt_endlednumber.Name = "txt_endlednumber";
            txt_endlednumber.Size = new Size(150, 27);
            txt_endlednumber.TabIndex = 4;
            txt_endlednumber.TextChanged += textClearOtherRange;
            // 
            // txt_endworkno
            // 
            txt_endworkno.BorderStyle = BorderStyle.FixedSingle;
            txt_endworkno.CharacterCasing = CharacterCasing.Upper;
            txt_endworkno.Font = new Font("Leelawadee UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            txt_endworkno.Location = new Point(278, 25);
            txt_endworkno.Name = "txt_endworkno";
            txt_endworkno.Size = new Size(150, 27);
            txt_endworkno.TabIndex = 2;
            txt_endworkno.TextChanged += textClearOtherRange;
            // 
            // txt_startlednumber
            // 
            txt_startlednumber.BorderStyle = BorderStyle.FixedSingle;
            txt_startlednumber.CharacterCasing = CharacterCasing.Upper;
            txt_startlednumber.Font = new Font("Leelawadee UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            txt_startlednumber.Location = new Point(90, 61);
            txt_startlednumber.Name = "txt_startlednumber";
            txt_startlednumber.Size = new Size(150, 27);
            txt_startlednumber.TabIndex = 3;
            txt_startlednumber.TextChanged += textClearOtherRange;
            // 
            // txt_startworkno
            // 
            txt_startworkno.BorderStyle = BorderStyle.FixedSingle;
            txt_startworkno.CharacterCasing = CharacterCasing.Upper;
            txt_startworkno.Font = new Font("Leelawadee UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            txt_startworkno.Location = new Point(90, 25);
            txt_startworkno.Name = "txt_startworkno";
            txt_startworkno.Size = new Size(150, 27);
            txt_startworkno.TabIndex = 1;
            txt_startworkno.TextChanged += textClearOtherRange;
            // 
            // tabPage3
            // 
            tabPage3.Controls.Add(groupBox4);
            tabPage3.Controls.Add(groupBox1);
            tabPage3.Location = new Point(4, 26);
            tabPage3.Name = "tabPage3";
            tabPage3.Padding = new Padding(3);
            tabPage3.Size = new Size(1000, 651);
            tabPage3.TabIndex = 2;
            tabPage3.Text = "Setting Data";
            tabPage3.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            groupBox4.Controls.Add(txt_maxmonth);
            groupBox4.Controls.Add(label30);
            groupBox4.Controls.Add(label24);
            groupBox4.Controls.Add(txt_firstInstalldate);
            groupBox4.Controls.Add(label23);
            groupBox4.Controls.Add(txt_dateatcalculate);
            groupBox4.Controls.Add(txt_datefest);
            groupBox4.Controls.Add(label31);
            groupBox4.Controls.Add(label29);
            groupBox4.Controls.Add(btn_save);
            groupBox4.Controls.Add(label22);
            groupBox4.Controls.Add(txt_festName);
            groupBox4.Controls.Add(label8);
            groupBox4.Controls.Add(label7);
            groupBox4.Controls.Add(txt_festNo);
            groupBox4.Controls.Add(label6);
            groupBox4.Location = new Point(23, 166);
            groupBox4.Name = "groupBox4";
            groupBox4.Size = new Size(951, 237);
            groupBox4.TabIndex = 6;
            groupBox4.TabStop = false;
            // 
            // txt_maxmonth
            // 
            txt_maxmonth.BorderStyle = BorderStyle.FixedSingle;
            txt_maxmonth.Location = new Point(568, 149);
            txt_maxmonth.Name = "txt_maxmonth";
            txt_maxmonth.Size = new Size(192, 25);
            txt_maxmonth.TabIndex = 7;
            txt_maxmonth.TextAlign = HorizontalAlignment.Center;
            // 
            // label30
            // 
            label30.AutoSize = true;
            label30.ForeColor = Color.Navy;
            label30.Location = new Point(400, 175);
            label30.Name = "label30";
            label30.Size = new Size(96, 17);
            label30.TabIndex = 6;
            label30.Text = "EX. 01/01/2025";
            // 
            // label24
            // 
            label24.AutoSize = true;
            label24.ForeColor = Color.Navy;
            label24.Location = new Point(247, 175);
            label24.Name = "label24";
            label24.Size = new Size(96, 17);
            label24.TabIndex = 6;
            label24.Text = "EX. 01/01/2025";
            // 
            // txt_firstInstalldate
            // 
            txt_firstInstalldate.BorderStyle = BorderStyle.FixedSingle;
            txt_firstInstalldate.Location = new Point(400, 147);
            txt_firstInstalldate.Mask = "00/00/0000";
            txt_firstInstalldate.Name = "txt_firstInstalldate";
            txt_firstInstalldate.Size = new Size(144, 25);
            txt_firstInstalldate.TabIndex = 5;
            txt_firstInstalldate.TextAlign = HorizontalAlignment.Center;
            txt_firstInstalldate.ValidatingType = typeof(DateTime);
            // 
            // label23
            // 
            label23.AutoSize = true;
            label23.ForeColor = Color.Navy;
            label23.Location = new Point(96, 175);
            label23.Name = "label23";
            label23.Size = new Size(96, 17);
            label23.TabIndex = 6;
            label23.Text = "EX. 01/01/2025";
            // 
            // txt_dateatcalculate
            // 
            txt_dateatcalculate.BorderStyle = BorderStyle.FixedSingle;
            txt_dateatcalculate.Location = new Point(247, 147);
            txt_dateatcalculate.Mask = "00/00/0000";
            txt_dateatcalculate.Name = "txt_dateatcalculate";
            txt_dateatcalculate.Size = new Size(144, 25);
            txt_dateatcalculate.TabIndex = 5;
            txt_dateatcalculate.TextAlign = HorizontalAlignment.Center;
            txt_dateatcalculate.ValidatingType = typeof(DateTime);
            // 
            // txt_datefest
            // 
            txt_datefest.BorderStyle = BorderStyle.FixedSingle;
            txt_datefest.Location = new Point(96, 147);
            txt_datefest.Mask = "00/00/0000";
            txt_datefest.Name = "txt_datefest";
            txt_datefest.Size = new Size(144, 25);
            txt_datefest.TabIndex = 5;
            txt_datefest.TextAlign = HorizontalAlignment.Center;
            txt_datefest.ValidatingType = typeof(DateTime);
            // 
            // label31
            // 
            label31.AutoSize = true;
            label31.Location = new Point(568, 129);
            label31.Name = "label31";
            label31.Size = new Size(192, 17);
            label31.TabIndex = 3;
            label31.Text = "งวดสุดท้าย ก่อนหมดอายุความ (เดือน)";
            // 
            // label29
            // 
            label29.AutoSize = true;
            label29.Location = new Point(400, 127);
            label29.Name = "label29";
            label29.Size = new Size(91, 17);
            label29.TabIndex = 3;
            label29.Text = "วันทีชำระงวดแรก";
            // 
            // btn_save
            // 
            btn_save.Image = (Image)resources.GetObject("btn_save.Image");
            btn_save.Location = new Point(835, 189);
            btn_save.Name = "btn_save";
            btn_save.Size = new Size(110, 33);
            btn_save.TabIndex = 4;
            btn_save.Text = "บันทึก";
            btn_save.TextAlign = ContentAlignment.MiddleRight;
            btn_save.TextImageRelation = TextImageRelation.ImageBeforeText;
            btn_save.UseVisualStyleBackColor = true;
            btn_save.Click += btn_save_Click;
            // 
            // label22
            // 
            label22.AutoSize = true;
            label22.Location = new Point(247, 127);
            label22.Name = "label22";
            label22.Size = new Size(90, 17);
            label22.TabIndex = 3;
            label22.Text = "ภาระหนี้ ณ วันที่";
            // 
            // txt_festName
            // 
            txt_festName.BorderStyle = BorderStyle.FixedSingle;
            txt_festName.Location = new Point(96, 95);
            txt_festName.Name = "txt_festName";
            txt_festName.Size = new Size(664, 25);
            txt_festName.TabIndex = 0;
            txt_festName.TextAlign = HorizontalAlignment.Center;
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Location = new Point(96, 127);
            label8.Name = "label8";
            label8.Size = new Size(63, 17);
            label8.TabIndex = 3;
            label8.Text = "วันที่จัดงาน";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(96, 75);
            label7.Name = "label7";
            label7.Size = new Size(77, 17);
            label7.TabIndex = 3;
            label7.Text = "สถานที่จัดงาน";
            // 
            // txt_festNo
            // 
            txt_festNo.BorderStyle = BorderStyle.FixedSingle;
            txt_festNo.Location = new Point(96, 40);
            txt_festNo.Name = "txt_festNo";
            txt_festNo.Size = new Size(183, 25);
            txt_festNo.TabIndex = 0;
            txt_festNo.TextAlign = HorizontalAlignment.Center;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(96, 20);
            label6.Name = "label6";
            label6.Size = new Size(180, 17);
            label6.TabIndex = 3;
            label6.Text = "มหกรรมไกล่เกลี่ยชั้นบังคับคดี ครั้งที่ ";
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(btn_table_cal);
            groupBox1.Controls.Add(btn_path_c2);
            groupBox1.Controls.Add(label2);
            groupBox1.Controls.Add(label1);
            groupBox1.Controls.Add(txt_path_table);
            groupBox1.Controls.Add(txt_path_c2);
            groupBox1.Location = new Point(23, 19);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(951, 141);
            groupBox1.TabIndex = 5;
            groupBox1.TabStop = false;
            groupBox1.Text = "Path เก็บเอกสาร";
            // 
            // btn_table_cal
            // 
            btn_table_cal.Location = new Point(766, 93);
            btn_table_cal.Name = "btn_table_cal";
            btn_table_cal.Size = new Size(34, 25);
            btn_table_cal.TabIndex = 4;
            btn_table_cal.Text = "....";
            btn_table_cal.UseVisualStyleBackColor = true;
            btn_table_cal.Click += btn_table_cal_Click;
            // 
            // btn_path_c2
            // 
            btn_path_c2.Location = new Point(766, 45);
            btn_path_c2.Name = "btn_path_c2";
            btn_path_c2.Size = new Size(34, 25);
            btn_path_c2.TabIndex = 4;
            btn_path_c2.Text = "....";
            btn_path_c2.UseVisualStyleBackColor = true;
            btn_path_c2.Click += btn_path_c2_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(96, 73);
            label2.Name = "label2";
            label2.Size = new Size(74, 17);
            label2.TabIndex = 3;
            label2.Text = "ตารางเจรจา :";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(96, 25);
            label1.Name = "label1";
            label1.Size = new Size(30, 17);
            label1.TabIndex = 3;
            label1.Text = "C2 :";
            // 
            // txt_path_table
            // 
            txt_path_table.BackColor = Color.White;
            txt_path_table.BorderStyle = BorderStyle.FixedSingle;
            txt_path_table.Location = new Point(96, 93);
            txt_path_table.Name = "txt_path_table";
            txt_path_table.ReadOnly = true;
            txt_path_table.Size = new Size(664, 25);
            txt_path_table.TabIndex = 3;
            // 
            // txt_path_c2
            // 
            txt_path_c2.BackColor = Color.White;
            txt_path_c2.BorderStyle = BorderStyle.FixedSingle;
            txt_path_c2.Location = new Point(96, 45);
            txt_path_c2.Name = "txt_path_c2";
            txt_path_c2.ReadOnly = true;
            txt_path_c2.Size = new Size(664, 25);
            txt_path_c2.TabIndex = 2;
            // 
            // frmReportManage
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1008, 681);
            Controls.Add(tabControl1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "frmReportManage";
            tabControl1.ResumeLayout(false);
            tabPage2.ResumeLayout(false);
            groupBox6.ResumeLayout(false);
            groupBox6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridShow).EndInit();
            groupBox5.ResumeLayout(false);
            groupBox5.PerformLayout();
            tabPage3.ResumeLayout(false);
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
        }

        #endregion
        private TabControl tabControl1;
        private TabPage tabPage2;
        private TabPage tabPage3;
        private GroupBox groupBox4;
        private GroupBox groupBox1;
        private Button btn_table_cal;
        private Button btn_path_c2;
        private Label label2;
        private Label label1;
        private TextBox txt_path_table;
        private TextBox txt_path_c2;
        private TextBox txt_festName;
        private Label label7;
        private TextBox txt_festNo;
        private Label label6;
        private Button btn_save;
        private Label label22;
        private Label label8;
        private MaskedTextBox txt_datefest;
        private Label label24;
        private Label label23;
        private MaskedTextBox txt_dateatcalculate;
        private GroupBox groupBox5;
        private Button brn_getwithrange;
        private Label label28;
        private Label label26;
        private Label label27;
        private Label label25;
        private TextBox txt_endlednumber;
        private TextBox txt_endworkno;
        private TextBox txt_startlednumber;
        private TextBox txt_startworkno;
        private GroupBox groupBox6;
        private DataGridView dataGridShow;
        private Button btn_C2TableRangePDF;
        private Button btn_TableRangePDF;
        private Button btn_C2RangePDF;
        private CheckBox chk_merge;
        private Label label30;
        private MaskedTextBox txt_firstInstalldate;
        private Label label29;
        private TextBox txt_maxmonth;
        private Label label31;
        private Button btn_export_Excel;
    }
}