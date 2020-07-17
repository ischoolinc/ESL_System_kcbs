using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using DevComponents.DotNetBar;

namespace ESL_System.Form
{
    public partial class TemplateReNameForm : BaseForm
    {
        ButtonItem _currentItem;

        public TemplateReNameForm(ButtonItem currentItem)
        {
            InitializeComponent();
            //取得目前所選的樣板
            _currentItem = currentItem;

            txtTemplateName.Text = currentItem.Text;
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtTemplateName.Text))
            {
                string esl_exam_template_id = "" + _currentItem.Tag;

                string new_esl_exam_template_name = txtTemplateName.Text;

                UpdateHelper uh = new UpdateHelper();

                //依照所選項目儲存
                string updQuery = "UPDATE exam_template SET name ='" + new_esl_exam_template_name + "' WHERE id ='" + esl_exam_template_id + "'";

                //執行sql，更新
                uh.Execute(updQuery);

                MsgBox.Show("樣板更名成功");

                DialogResult = DialogResult.OK;
            }
            else
            {
                MsgBox.Show("請輸入樣板名稱");
            }            
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}
