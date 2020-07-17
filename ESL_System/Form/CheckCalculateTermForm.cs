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
using FISCA.Data;

namespace ESL_System.Form
{
    public partial class CheckCalculateTermForm : BaseForm
    {
        // 本次 ESL 成績要換算成 對應系統的 試別 id
        private string _TargetExamId = "";

        private Dictionary<string, string> _ExamDict = new Dictionary<string, string>();
        private List<string> _CourseIDList;
        private Dictionary<string, List<string>> _CourseExamIDListDict = new Dictionary<string, List<string>>(); // 作為檢查該課程的評分樣板是否有該評量 <course_name,List<examID>>


        public CheckCalculateTermForm(List<string> courseIDList)
        {
            InitializeComponent();
            _CourseIDList = courseIDList;

        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            _TargetExamId = _ExamDict[comboBoxEx1.Text];

            if (_TargetExamId == "")
            {
                MsgBox.Show("請選擇計算試別!");
                return;
            }

            string courseIDs = string.Join(",", _CourseIDList);

            // 抓取課程上樣板，是否真的有設定所選對應試別
            string query = @"
SELECT 
	course.id AS course_id
    ,course.course_name
	,te_include.ref_exam_id 
	,exam_name
FROM course 	
	LEFT JOIN te_include ON te_include.ref_exam_template_id = course.ref_exam_template_id
	LEFT JOIN exam ON exam.id = te_include.ref_exam_id
WHERE course.id IN ( " + courseIDs + @")
ORDER BY course.id,ref_exam_id";

            
            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query);

            _CourseExamIDListDict.Clear(); // 清空

            //整理目前的ESL 課程資料，其評分樣板有的評量
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (!_CourseExamIDListDict.ContainsKey("" + dr["course_name"]))
                    {
                        _CourseExamIDListDict.Add("" + dr["course_name"], new List<string>());
                        _CourseExamIDListDict["" + dr["course_name"]].Add("" + dr["ref_exam_id"]);

                    }
                    else
                    {
                        _CourseExamIDListDict["" + dr["course_name"]].Add("" + dr["ref_exam_id"]);
                    }
                }
            }

            List<string> errorList = new List<string>();

            foreach (KeyValuePair<string, List<string>> p in _CourseExamIDListDict)
            {
                if (!p.Value.Contains(_TargetExamId))
                {
                    errorList.Add("所選取課程:「" + p.Key + "」，其評分樣本上無設定試別:「" + comboBoxEx1.Text + "」，請至教務作業/評分樣板設定 修改。");
                }
            }

            if (errorList.Count > 0)
            {
                string erroor = string.Join("\r\n", errorList);

                MsgBox.Show(erroor);

                return; 
            }


            CalculateTermExamScore ctes = new CalculateTermExamScore(_CourseIDList, _TargetExamId);

            ctes.CalculateESLTermScore();

            this.DialogResult = DialogResult.Yes;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void CheckCalculateTermForm_Load(object sender, EventArgs e)
        {
            string courseIDs = string.Join(",", _CourseIDList);

            // 若使用者選取課程沒有 ESL 的樣板設定 則提醒
            string query = @"
SELECT 
    course.id
    ,course.course_name
    ,exam_template.description 
FROM course 
LEFT JOIN  exam_template ON course.ref_exam_template_id =exam_template.id  
WHERE course.id IN( " + courseIDs + ") AND  exam_template.description IS NULL  ";

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query);
            if (dt.Rows.Count > 0)
            {
                List<string> errorList = new List<string>();

                foreach (DataRow dr in dt.Rows)
                {
                    errorList.Add("所選取課程:「" + dr["course_name"] + "」，其並非使用ESL評分樣版，無法使用本功能。 ");
                }

                if (errorList.Count > 0)
                {
                    string erroor = string.Join("\r\n", errorList);

                    MsgBox.Show(erroor);

                    this.Close();
                }
            }
                                 
            // 抓取系統內 設定的試別
            query = "SELECT id,exam_name FROM exam ";

            qh = new QueryHelper();
            dt = qh.Select(query);

            _ExamDict.Clear(); // 清空

            //整理目前的ESL 課程資料
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    _ExamDict.Add("" + dr["exam_name"], "" + dr["id"]); // <name,id>
                }
            }
            else
            {
                MsgBox.Show("系統內無任何評量設定，請至教務作業設定。");
                this.Close();
            }


            foreach (KeyValuePair<string, string> exam in _ExamDict)
            {
                object o = exam.Key;
                
                comboBoxEx1.Items.Add(o);
               
            }
        }
    }
}
