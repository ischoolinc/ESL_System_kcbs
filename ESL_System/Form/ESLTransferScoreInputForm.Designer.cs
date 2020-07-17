namespace ESL_System.Form
{
    partial class ESLTransferScoreInputForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridViewX1 = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.ColClass = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColSeatNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColStudentNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColTerm = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColSubject = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColAssessment = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColTeacher = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColScore = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColRatio = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.buttonX1 = new DevComponents.DotNetBar.ButtonX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.cboExam = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.picLoading = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewX1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picLoading)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewX1
            // 
            this.dataGridViewX1.AllowUserToAddRows = false;
            this.dataGridViewX1.AllowUserToDeleteRows = false;
            this.dataGridViewX1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewX1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewX1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColClass,
            this.ColSeatNo,
            this.ColName,
            this.ColStudentNumber,
            this.ColTerm,
            this.ColSubject,
            this.ColAssessment,
            this.ColTeacher,
            this.ColScore,
            this.ColRatio});
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewX1.DefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridViewX1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridViewX1.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dataGridViewX1.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.dataGridViewX1.Location = new System.Drawing.Point(12, 84);
            this.dataGridViewX1.MultiSelect = false;
            this.dataGridViewX1.Name = "dataGridViewX1";
            this.dataGridViewX1.RowTemplate.Height = 24;
            this.dataGridViewX1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewX1.Size = new System.Drawing.Size(1228, 560);
            this.dataGridViewX1.TabIndex = 0;
            this.dataGridViewX1.CellValidated += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewX1_CellValidated);
            this.dataGridViewX1.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dataGridViewX1_CellValidating);
            // 
            // ColClass
            // 
            this.ColClass.HeaderText = "班級";
            this.ColClass.Name = "ColClass";
            this.ColClass.ReadOnly = true;
            // 
            // ColSeatNo
            // 
            this.ColSeatNo.HeaderText = "座號";
            this.ColSeatNo.Name = "ColSeatNo";
            this.ColSeatNo.ReadOnly = true;
            // 
            // ColName
            // 
            this.ColName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ColName.HeaderText = "姓名";
            this.ColName.Name = "ColName";
            this.ColName.ReadOnly = true;
            // 
            // ColStudentNumber
            // 
            this.ColStudentNumber.HeaderText = "學號";
            this.ColStudentNumber.Name = "ColStudentNumber";
            this.ColStudentNumber.ReadOnly = true;
            // 
            // ColTerm
            // 
            this.ColTerm.HeaderText = "試別";
            this.ColTerm.Name = "ColTerm";
            // 
            // ColSubject
            // 
            this.ColSubject.HeaderText = "科目";
            this.ColSubject.Name = "ColSubject";
            // 
            // ColAssessment
            // 
            this.ColAssessment.HeaderText = "評量";
            this.ColAssessment.Name = "ColAssessment";
            // 
            // ColTeacher
            // 
            this.ColTeacher.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ColTeacher.HeaderText = "教師";
            this.ColTeacher.Name = "ColTeacher";
            this.ColTeacher.ReadOnly = true;
            // 
            // ColScore
            // 
            this.ColScore.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ColScore.HeaderText = "分數";
            this.ColScore.Name = "ColScore";
            // 
            // ColRatio
            // 
            this.ColRatio.HeaderText = "比重";
            this.ColRatio.Name = "ColRatio";
            // 
            // buttonX1
            // 
            this.buttonX1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonX1.BackColor = System.Drawing.Color.Transparent;
            this.buttonX1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.buttonX1.Location = new System.Drawing.Point(1165, 652);
            this.buttonX1.Name = "buttonX1";
            this.buttonX1.Size = new System.Drawing.Size(75, 23);
            this.buttonX1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonX1.TabIndex = 1;
            this.buttonX1.Text = "儲存";
            this.buttonX1.Click += new System.EventHandler(this.buttonX1_Click);
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(12, 12);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(34, 21);
            this.labelX1.TabIndex = 3;
            this.labelX1.Text = "課程";
            // 
            // labelX2
            // 
            this.labelX2.AutoSize = true;
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(770, 44);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(470, 21);
            this.labelX2.TabIndex = 4;
            this.labelX2.Text = "若人工指定成績比重，本項目計算時將只以人工比例計算，而非ESL樣板設定。";
            // 
            // cboExam
            // 
            this.cboExam.DisplayMember = "Text";
            this.cboExam.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboExam.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboExam.Enabled = false;
            this.cboExam.FormattingEnabled = true;
            this.cboExam.ItemHeight = 19;
            this.cboExam.Location = new System.Drawing.Point(90, 40);
            this.cboExam.Name = "cboExam";
            this.cboExam.Size = new System.Drawing.Size(231, 25);
            this.cboExam.TabIndex = 5;
            this.cboExam.SelectedIndexChanged += new System.EventHandler(this.cboExam_SelectedIndexChanged);
            // 
            // labelX3
            // 
            this.labelX3.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.Class = "";
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Location = new System.Drawing.Point(12, 40);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(75, 23);
            this.labelX3.TabIndex = 6;
            this.labelX3.Text = "請選擇試別";
            // 
            // picLoading
            // 
            this.picLoading.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.picLoading.BackColor = System.Drawing.Color.Transparent;
            this.picLoading.Image = global::ESL_System.Properties.Resources.loading;
            this.picLoading.InitialImage = null;
            this.picLoading.Location = new System.Drawing.Point(604, 320);
            this.picLoading.Name = "picLoading";
            this.picLoading.Size = new System.Drawing.Size(44, 46);
            this.picLoading.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picLoading.TabIndex = 11;
            this.picLoading.TabStop = false;
            this.picLoading.Visible = false;
            // 
            // ESLTransferScoreInputForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1252, 687);
            this.Controls.Add(this.picLoading);
            this.Controls.Add(this.cboExam);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.buttonX1);
            this.Controls.Add(this.dataGridViewX1);
            this.DoubleBuffered = true;
            this.MaximumSize = new System.Drawing.Size(1268, 726);
            this.MinimumSize = new System.Drawing.Size(1268, 726);
            this.Name = "ESLTransferScoreInputForm";
            this.Text = "ESL課程缺考成績處理";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewX1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picLoading)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX dataGridViewX1;
        private DevComponents.DotNetBar.ButtonX buttonX1;
        private DevComponents.DotNetBar.LabelX labelX1;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColClass;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColSeatNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColStudentNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColTerm;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColSubject;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColAssessment;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColTeacher;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColScore;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColRatio;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboExam;
        private DevComponents.DotNetBar.LabelX labelX3;
        private System.Windows.Forms.PictureBox picLoading;
    }
}