using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace ESL_System.Form
{
    public partial class NewConfigure : FISCA.Presentation.Controls.BaseForm
    {
        public Aspose.Words.Document Template { get; private set; }
        public int SubjectLimit { get; private set; }
        public int AttendanceCountLimit  { get; private set; }
        public int AttendanceDetailLimit  { get; private set; }
        public int DisciplineDetailLimit  { get; private set; }
        public int ServiceLearningDetailLimit { get; private set; }


        public string ConfigName { get; private set; }

        public NewConfigure()
        {
            InitializeComponent();            
            checkBoxX2.CheckedChanged += new EventHandler(UploadTemplate);
        }

        private void UploadTemplate(object sender, EventArgs e)
        {
            if (checkBoxX2.Checked)
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Title = "上傳樣板";
                dialog.Filter = "Word檔案 (*.doc)|*.doc|Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        Template = new Aspose.Words.Document(dialog.FileName);
                                                                  
                    }
                    catch
                    {
                        MessageBox.Show("樣板開啟失敗");
                        checkBoxX2.Checked = false;
                    }
                }
                else
                    checkBoxX2.Checked = false;
            }
        }

        private void checkReady(object sender, EventArgs e)
        {
            bool ready = true;
            if (txtName.Text == "")
                ready = false;
            else
                ConfigName = txtName.Text;
            if (!checkBoxX2.Checked)
            {
                ready = false;
            }
            btnSubmit.Enabled = ready;
        }

        
        // 下載合併欄位總表
        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtName.Text))
                return;
           
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            Close();
        }

        private void lnkMore_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            
        }

    }
}
