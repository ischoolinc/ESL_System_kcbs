namespace ESL_System.Form
{
    partial class ReportTemplateSettingForm
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
            this.labelX15 = new DevComponents.DotNetBar.LabelX();
            this.labelX14 = new DevComponents.DotNetBar.LabelX();
            this.labelX13 = new DevComponents.DotNetBar.LabelX();
            this.linkLabel5 = new System.Windows.Forms.LinkLabel();
            this.lbv01 = new System.Windows.Forms.LinkLabel();
            this.lbc01 = new System.Windows.Forms.LinkLabel();
            this.comboBoxEx1 = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.comboBoxEx2 = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.buttonX2 = new DevComponents.DotNetBar.ButtonX();
            this.SuspendLayout();
            // 
            // labelX15
            // 
            this.labelX15.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.labelX15.AutoSize = true;
            this.labelX15.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX15.BackgroundStyle.Class = "";
            this.labelX15.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX15.Location = new System.Drawing.Point(7, 39);
            this.labelX15.Name = "labelX15";
            this.labelX15.Size = new System.Drawing.Size(74, 21);
            this.labelX15.TabIndex = 34;
            this.labelX15.Text = "樣板設定：";
            // 
            // labelX14
            // 
            this.labelX14.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.labelX14.AutoSize = true;
            this.labelX14.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX14.BackgroundStyle.Class = "";
            this.labelX14.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX14.Location = new System.Drawing.Point(179, 11);
            this.labelX14.Name = "labelX14";
            this.labelX14.Size = new System.Drawing.Size(47, 21);
            this.labelX14.TabIndex = 31;
            this.labelX14.Text = "學期：";
            // 
            // labelX13
            // 
            this.labelX13.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.labelX13.AutoSize = true;
            this.labelX13.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX13.BackgroundStyle.Class = "";
            this.labelX13.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX13.Location = new System.Drawing.Point(7, 11);
            this.labelX13.Name = "labelX13";
            this.labelX13.Size = new System.Drawing.Size(60, 21);
            this.labelX13.TabIndex = 30;
            this.labelX13.Text = "學年度：";
            // 
            // linkLabel5
            // 
            this.linkLabel5.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.linkLabel5.AutoSize = true;
            this.linkLabel5.BackColor = System.Drawing.Color.Transparent;
            this.linkLabel5.Location = new System.Drawing.Point(223, 63);
            this.linkLabel5.Name = "linkLabel5";
            this.linkLabel5.Size = new System.Drawing.Size(112, 17);
            this.linkLabel5.TabIndex = 28;
            this.linkLabel5.TabStop = true;
            this.linkLabel5.Text = "下載合併欄位總表";
            this.linkLabel5.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel5_LinkClicked);
            // 
            // lbv01
            // 
            this.lbv01.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.lbv01.AutoSize = true;
            this.lbv01.BackColor = System.Drawing.Color.Transparent;
            this.lbv01.Location = new System.Drawing.Point(4, 63);
            this.lbv01.Name = "lbv01";
            this.lbv01.Size = new System.Drawing.Size(86, 17);
            this.lbv01.TabIndex = 26;
            this.lbv01.TabStop = true;
            this.lbv01.Text = "檢視套印樣板";
            this.lbv01.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lbv01_LinkClicked);
            // 
            // lbc01
            // 
            this.lbc01.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.lbc01.AutoSize = true;
            this.lbc01.BackColor = System.Drawing.Color.Transparent;
            this.lbc01.Location = new System.Drawing.Point(114, 63);
            this.lbc01.Name = "lbc01";
            this.lbc01.Size = new System.Drawing.Size(86, 17);
            this.lbc01.TabIndex = 27;
            this.lbc01.TabStop = true;
            this.lbc01.Text = "變更套印樣板";
            this.lbc01.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lbc01_LinkClicked);
            // 
            // comboBoxEx1
            // 
            this.comboBoxEx1.DisplayMember = "Text";
            this.comboBoxEx1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBoxEx1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxEx1.FormattingEnabled = true;
            this.comboBoxEx1.ItemHeight = 19;
            this.comboBoxEx1.Location = new System.Drawing.Point(66, 8);
            this.comboBoxEx1.Name = "comboBoxEx1";
            this.comboBoxEx1.Size = new System.Drawing.Size(107, 25);
            this.comboBoxEx1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.comboBoxEx1.TabIndex = 35;
            // 
            // comboBoxEx2
            // 
            this.comboBoxEx2.DisplayMember = "Text";
            this.comboBoxEx2.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBoxEx2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxEx2.FormattingEnabled = true;
            this.comboBoxEx2.ItemHeight = 19;
            this.comboBoxEx2.Location = new System.Drawing.Point(226, 8);
            this.comboBoxEx2.Name = "comboBoxEx2";
            this.comboBoxEx2.Size = new System.Drawing.Size(109, 25);
            this.comboBoxEx2.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.comboBoxEx2.TabIndex = 36;
            // 
            // buttonX2
            // 
            this.buttonX2.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX2.BackColor = System.Drawing.Color.Transparent;
            this.buttonX2.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.buttonX2.Location = new System.Drawing.Point(261, 90);
            this.buttonX2.Name = "buttonX2";
            this.buttonX2.Size = new System.Drawing.Size(75, 23);
            this.buttonX2.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonX2.TabIndex = 38;
            this.buttonX2.Text = "離開";
            this.buttonX2.Click += new System.EventHandler(this.buttonX2_Click);
            // 
            // ReportTemplateSettingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(348, 117);
            this.Controls.Add(this.buttonX2);
            this.Controls.Add(this.comboBoxEx2);
            this.Controls.Add(this.comboBoxEx1);
            this.Controls.Add(this.labelX15);
            this.Controls.Add(this.labelX14);
            this.Controls.Add(this.labelX13);
            this.Controls.Add(this.linkLabel5);
            this.Controls.Add(this.lbv01);
            this.Controls.Add(this.lbc01);
            this.DoubleBuffered = true;
            this.MaximumSize = new System.Drawing.Size(364, 156);
            this.MinimumSize = new System.Drawing.Size(364, 156);
            this.Name = "ReportTemplateSettingForm";
            this.Text = "樣板設定";
            this.Load += new System.EventHandler(this.ReportTemplateSettingForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelX15;
        private DevComponents.DotNetBar.LabelX labelX14;
        private DevComponents.DotNetBar.LabelX labelX13;
        private System.Windows.Forms.LinkLabel linkLabel5;
        private System.Windows.Forms.LinkLabel lbv01;
        private System.Windows.Forms.LinkLabel lbc01;
        private DevComponents.DotNetBar.Controls.ComboBoxEx comboBoxEx1;
        private DevComponents.DotNetBar.Controls.ComboBoxEx comboBoxEx2;
        private DevComponents.DotNetBar.ButtonX buttonX2;
    }
}