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
using FISCA.LogAgent;
using FISCA.Authentication;


namespace ESL_System.Form
{
    public partial class ESLScoreInputForm : FISCA.Presentation.Controls.BaseForm
    {
        // 目標課程名稱
        private string _targetCourseName;

        // 目標評量名稱
        private string _targetTermName;

        // 目標科目名稱
        private string _targetSubjectName;

        // 目標評量名稱
        private string _targetAssessmentName;

        // 目標分數種類(Score、Indicator、Comment)
        private string _targetScoreType;

        // 指標性評語 可以使用項目
        private List<Indicators> _targetIndicatorList = new List<Indicators>();

        // 評語型項目 字數上限
        private string _commentLimit;

        // 目前本課程的 ESL 評分樣版
        private ESLTemplate _eslTemplate;

        // ESL 分數 <subjectName,List<ESL分數>>
        private Dictionary<string, List<ESLScore>> _scoreDict = new Dictionary<string, List<ESLScore>>();

        // 修課學生 List
        private List<K12.Data.SCAttendRecord> _scaList;

        private BackgroundWorker _worker;
      
        // 新版會依照 使用者點選的評分項目 直接顯示對應的分數項目
        public ESLScoreInputForm(string targetCourseName, ESLTemplate eslTemplate, Dictionary<string, List<ESLScore>> scoreDict, string targetTermName, string targetSubjectName, string targetAssessmentName, List<K12.Data.SCAttendRecord> scaList)
        {
            InitializeComponent();

            _targetCourseName = targetCourseName;
            _targetTermName = targetTermName;
            _targetSubjectName = targetSubjectName;
            _targetAssessmentName = targetAssessmentName;

            _eslTemplate = eslTemplate;

            // 對出 目前的 評量的 分數類型
            _targetScoreType = _eslTemplate.TermList.Find(t => t.Name == _targetTermName).SubjectList.Find(s => s.Name == _targetSubjectName).AssessmentList.Find(a => a.Name == _targetAssessmentName).Type;

            // 對出 目前的 指標型分數 可以用的 指標
            _targetIndicatorList = _eslTemplate.TermList.Find(t => t.Name == _targetTermName).SubjectList.Find(s => s.Name == _targetSubjectName).AssessmentList.Find(a => a.Name == _targetAssessmentName).IndicatorsList;
             
            // 對出 目前的 評語型分數 可以輸入的字數上限
            _commentLimit = _eslTemplate.TermList.Find(t => t.Name == _targetTermName).SubjectList.Find(s => s.Name == _targetSubjectName).AssessmentList.Find(a => a.Name == _targetAssessmentName).InputLimit;

            _scoreDict = scoreDict;

            _scaList = scaList;

            labelX1.Text = _targetCourseName + " / " + _targetSubjectName + " / " + _targetAssessmentName; // 課程名稱 / 科目名稱 / 評量名稱

            // 不同的分數型態 有不同的表頭
            switch (_targetScoreType)
            {
                case "Score":
                    ColScore.HeaderText = "分數";
                    break;
                case "Indicator":
                    ColScore.HeaderText = "指標";
                    break;
                case "Comment":
                    ColScore.HeaderText = "評語";
                    break;
                default:
                    ColScore.HeaderText = "分數";
                    break;
            }

            // 填入分數
            FillScore();
        }







        private void FillScore()
        {
            dataGridViewX1.Rows.Clear();

           
            List<ESLScore> eslScoreList = new List<ESLScore>();

            foreach (ESLScore scoreItem in _scoreDict[_targetSubjectName])
            {
                if (scoreItem.Assessment == _targetAssessmentName)
                {
                    DataGridViewRow row = new DataGridViewRow();

                    row.CreateCells(dataGridViewX1);

                    K12.Data.SCAttendRecord scar = _scaList.Find(sca => sca.RefStudentID == scoreItem.RefStudentID && sca.RefCourseID == scoreItem.RefCourseID);

                    row.Cells[0].Value = scar.Student.Class != null ? scar.Student.Class.Name : ""; // 學生班級
                    row.Cells[1].Value = scar.Student != null ? "" + scar.Student.SeatNo : "";  // 學生座號
                    row.Cells[2].Value = scar.Student != null ? "" + scar.Student.Name : "";      // 學生姓名
                    row.Cells[3].Value = scar.Student != null ? "" + scar.Student.StudentNumber : "";  // 學生學號
                    row.Cells[4].Value = scoreItem.RefTeacherName;  // 本項目老師姓名
                    row.Cells[5].Value = scoreItem.HasValue ? scoreItem.Value : "";  // 本項目分數

                    row.Tag =  scoreItem.RefStudentID;  // row tag 用studentID 就夠

                    dataGridViewX1.Rows.Add(row);

                }
            }

            // 依   學號 排序 (同Web 的成績輸入介面)
            dataGridViewX1.Sort(ColStudentNumber, ListSortDirection.Ascending);

        }


        // 儲存
        private void buttonX1_Click(object sender, EventArgs e)
        {
            // 輸入分數 超過 界線提醒
            bool outRangeWarning = false;

            // 檢查是否有老師
            bool hasNoTeacher = false;

            if (_targetScoreType == "Score")
            {
                decimal i = 0;

                foreach (DataGridViewRow row in dataGridViewX1.Rows)
                {
                    //只檢查分數欄， 其值 若 超出 0~100 的範圍 跳出提醒視窗 是否繼續儲存
                    DataGridViewCell cell = dataGridViewX1.Rows[row.Index].Cells[5];

                    if (decimal.TryParse("" + cell.Value, out i))
                    {
                        if (i < 0 || i > 100)
                        {
                            outRangeWarning = true;
                        }
                    }

                    // 沒有老師
                    if ("" + dataGridViewX1.Rows[row.Index].Cells[4].Value == "")
                    {
                        hasNoTeacher = true;
                    }
                }
            }

            // 沒有老師 不給存
            if (hasNoTeacher)
            {
                MsgBox.Show("此成績項目沒有老師，請至課程指定教師。", "錯誤!", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                return;                
            }

            // 若畫面上有分數超過範圍，則跳出提醒視窗， 讓使用者決定是否要繼續儲存。
            if (outRangeWarning)
            {
                if (MsgBox.Show("表格上有分數成績不在0~100範圍內，請問是否繼續儲存?", "警告!", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                {
                    return;
                }
            }
            
            _worker = new BackgroundWorker();
            _worker.DoWork += new DoWorkEventHandler(Worker_DoWork);
            _worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Worker_RunWorkerCompleted);
            _worker.ProgressChanged += new ProgressChangedEventHandler(Worker_ProgressChanged);
            _worker.WorkerReportsProgress = true;

            // 暫停畫面控制項
            dataGridViewX1.SuspendLayout();
            buttonX1.Enabled = false;
           
            _worker.RunWorkerAsync();
        }



        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            _worker.ReportProgress(30,"整理成績資料...");

            //拚SQL
            // 兜資料
            List<string> dataList = new List<string>();

            List<ESLScore> updateESLscoreList = new List<ESLScore>(); // 最後要update ESLscoreList
            List<ESLScore> insertESLscoreList = new List<ESLScore>(); // 最後要indert ESLscoreList

            int i = 0;

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                foreach (ESLScore scoreItem in _scoreDict[_targetSubjectName])
                {
                    // 評量相同 、 學生ID 相同 則為該成績
                    if (scoreItem.Assessment == _targetAssessmentName && scoreItem.RefStudentID == "" + row.Tag)
                    {
                        // 本來有成績， 但值不同了 加入 更新名單
                        if (scoreItem.HasValue && scoreItem.Value != "" + row.Cells[5].Value)
                        {
                            scoreItem.OldValue = ""+ scoreItem.Value; // 舊分數

                            scoreItem.Value = "" + row.Cells[5].Value; // 新分數
                            
                            scoreItem.Value = scoreItem.Value.Trim().Replace("'", "''"); // trim 掉空白、 單引號特殊字

                            updateESLscoreList.Add(scoreItem);
                        }

                        // 本來沒成績， 本次新填的值 ，加入新增名單
                        if (!scoreItem.HasValue && "" + row.Cells[5].Value != "")
                        {
                            scoreItem.Value = "" + row.Cells[5].Value; // 新分數

                            scoreItem.Value = scoreItem.Value.Trim().Replace("'", "''"); // trim 掉空白、 單引號特殊字

                            scoreItem.HasValue = true;

                            insertESLscoreList.Add(scoreItem);
                        }
                    }
                }

                _worker.ReportProgress(30 + (i++ * 60 / dataGridViewX1.Rows.Count) , "整理成績資料...");

            }

            foreach (ESLScore score in updateESLscoreList)
            {
                string data = string.Format(@"
                SELECT
                    '{0}'::BIGINT AS ref_student_id
                    ,'{1}'::BIGINT AS ref_course_id
                    ,'{2}'::BIGINT AS ref_sc_attend_id
                    ,'{3}'::BIGINT AS ref_teacher_id
                    ,'{4}'::TEXT AS term
                    ,'{5}'::TEXT AS subject
                    ,'{6}'::TEXT AS assessment
                    ,'{7}'::TEXT AS custom_assessment
                    ,'{8}'::TEXT AS value
                    ,'{9}'::TEXT AS old_value
                    ,'{10}'::INTEGER AS uid
                    ,'UPDATE'::TEXT AS action
                ", score.RefStudentID, score.RefCourseID, score.RefScAttendID, score.RefTeacherID, score.Term, score.Subject,score.Assessment,"",score.Value, score.OldValue, score.ID);

                dataList.Add(data);
            }

            foreach (ESLScore score in insertESLscoreList)
            {
                string data = string.Format(@"
                SELECT
                    '{0}'::BIGINT AS ref_student_id
                    ,'{1}'::BIGINT AS ref_course_id
                    ,'{2}'::BIGINT AS ref_sc_attend_id
                    ,'{3}'::BIGINT AS ref_teacher_id
                    ,'{4}'::TEXT AS term
                    ,'{5}'::TEXT AS subject
                    ,'{6}'::TEXT AS assessment
                    ,'{7}'::TEXT AS custom_assessment
                    ,'{8}'::TEXT AS value
                    ,'{9}'::TEXT AS old_value
                    ,'{10}'::INTEGER AS uid
                    ,'INSERT'::TEXT AS action
                ", score.RefStudentID, score.RefCourseID, score.RefScAttendID, score.RefTeacherID, score.Term, score.Subject,score.Assessment,"",score.Value,"", 0);  // insert 給 oldValue ""、 uid = 0

                dataList.Add(data);
            }

            if (dataList.Count == 0)
            {
                return;
            }


            string Data = string.Join(" UNION ALL", dataList);

            // LOG 資訊
            string _actor = DSAServices.UserAccount; ;
            string _client_info = ClientInfo.GetCurrentClientInfo().OutputResult().OuterXml;


            string sql = string.Format(@"
WITH score_data_row AS(			 
                {0}     
),score_data_row_with_log_data AS(
	SELECT 
	score_data_row.ref_student_id
	,score_data_row.ref_course_id
	,score_data_row.ref_sc_attend_id
	,score_data_row.ref_teacher_id
	,score_data_row.term
	,score_data_row.subject
	,score_data_row.assessment
	,score_data_row.custom_assessment
	,score_data_row.value
	,score_data_row.old_value
	,score_data_row.uid
	,score_data_row.action	
	,student.name AS student_name	
	,student.student_number AS student_number	
	,course.course_name AS course_name	
	FROM score_data_row 
	LEFT JOIN student ON score_data_row.ref_student_id = student.id
	LEFT JOIN course ON score_data_row.ref_course_id = course.id
),update_score AS(	    
    Update $esl.gradebook_assessment_score
    SET
        ref_student_id = score_data_row.ref_student_id
        ,ref_course_id = score_data_row.ref_course_id
        ,ref_sc_attend_id = score_data_row.ref_sc_attend_id
        ,ref_teacher_id = score_data_row.ref_teacher_id
        ,term = score_data_row.term
        ,subject = score_data_row.subject
        ,assessment = score_data_row.assessment
        ,custom_assessment = score_data_row.custom_assessment
        ,value = score_data_row.value
        ,last_update = NOW()
    FROM 
        score_data_row    
    WHERE $esl.gradebook_assessment_score.uid = score_data_row.uid  
        AND score_data_row.action ='UPDATE'
    --RETURNING  $esl.gradebook_assessment_score.* 
),insert_score AS (
INSERT INTO $esl.gradebook_assessment_score(
	ref_student_id	
	,ref_course_id
    ,ref_sc_attend_id
	,ref_teacher_id
	,term
	,subject
    ,assessment
    ,custom_assessment
	,value
)
SELECT 
	score_data_row.ref_student_id::BIGINT AS ref_student_id	
	,score_data_row.ref_course_id::BIGINT AS ref_course_id	
    ,score_data_row.ref_sc_attend_id::BIGINT AS ref_sc_attend_id	
	,score_data_row.ref_teacher_id::BIGINT AS ref_teacher_id	
	,score_data_row.term::TEXT AS term	
	,score_data_row.subject::TEXT AS subject	
    ,score_data_row.assessment::TEXT AS assessment	
    ,score_data_row.custom_assessment::TEXT AS custom_assessment
	,score_data_row.value::TEXT AS value	
FROM
	score_data_row
WHERE action ='INSERT'
),insert_log AS(
INSERT INTO log(
	actor
	, action_type
	, action
	, target_category
	, target_id
	, server_time
	, client_info
	, action_by
	, description
)
SELECT 
	'{1}'::TEXT AS actor
	, 'Record' AS action_type
	, 'ESL後端成績輸入' AS action
	, 'student'::TEXT AS target_category
	, score_data_row_with_log_data.ref_student_id AS target_id
	, now() AS server_time
	, '{2}' AS client_info
	, 'ESL後端成績輸入'AS action_by   
	, '學號「'|| score_data_row_with_log_data.student_number||'」，學生「'|| score_data_row_with_log_data.student_name ||'」，ESL課程「'|| score_data_row_with_log_data.course_name || '」，試別「'||score_data_row_with_log_data.term|| '」，科目 「'||score_data_row_with_log_data.subject|| '」，評量「'||score_data_row_with_log_data.assessment|| '」，新增成績「'||score_data_row_with_log_data.value|| '」  ' AS description 
FROM
	score_data_row_with_log_data
WHERE action ='INSERT'
)
INSERT INTO log(
	actor
	, action_type
	, action
	, target_category
	, target_id
	, server_time
	, client_info
	, action_by
	, description
)
SELECT 
	'{1}'::TEXT AS actor
	, 'Record' AS action_type
	, 'ESL後端成績輸入' AS action
	, 'student'::TEXT AS target_category
	, score_data_row_with_log_data.ref_student_id AS target_id
	, now() AS server_time
	, '{2}' AS client_info
	, 'ESL後端成績輸入'AS action_by   
	, '學號「'|| score_data_row_with_log_data.student_number||'」，學生「'|| score_data_row_with_log_data.student_name ||'」，ESL課程「'|| score_data_row_with_log_data.course_name || '」，試別「'||score_data_row_with_log_data.term|| '」，科目 「'||score_data_row_with_log_data.subject|| '」，評量「'||score_data_row_with_log_data.assessment|| '」，成績自「'||score_data_row_with_log_data.old_value|| '」修改為「'||score_data_row_with_log_data.value|| '」  ' AS description 
FROM
	score_data_row_with_log_data
WHERE action ='UPDATE'
", Data,_actor,_client_info);



            UpdateHelper uh = new UpdateHelper();

            _worker.ReportProgress(90, "上傳成績...");

            //執行sql
            try
            {
                uh.Execute(sql);
                _worker.ReportProgress(100, "ESL 評量成績上傳完成。");
            }
            catch(Exception ex)
            {
                throw ex;
            }
                       
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            dataGridViewX1.ResumeLayout();            
            buttonX1.Enabled = true;

            if (e.Error != null)
            {
                MsgBox.Show("上傳失敗!!，錯誤訊息:" + e.Error.Message);

            }
            else if (e.Cancelled)
            {
                MsgBox.Show("上傳中止!!");
            }
            else
            {
                MsgBox.Show("上傳完成!");
            }
            
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage(""+e.UserState, e.ProgressPercentage);
        }

        private void dataGridViewX1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            // 只驗第五格 分數、指標、評語 欄位
            if (e.ColumnIndex != 5)
            {
                return;
            }

            DataGridViewCell cell = dataGridViewX1.Rows[e.RowIndex].Cells[e.ColumnIndex];

            cell.ErrorText = String.Empty;

            if (_targetScoreType == "Score")
            {
                decimal i = 0;

                if (!decimal.TryParse("" + e.FormattedValue, out i))
                {
                    cell.ErrorText = "請輸入數值。";                    
                }
            }
            if (_targetScoreType == "Indicator")
            {
                if (_targetIndicatorList.Find(indicator => indicator.Name == "" + e.FormattedValue) ==null)
                {
                    List<string> indicatorList = new List<string>();

                    string indicators = "";

                    foreach (Indicators indicator in _targetIndicatorList)
                    {
                        indicatorList.Add(indicator.Name);
                    }

                    indicators = string.Join("、", indicatorList);

                    cell.ErrorText = "請輸入" + indicators + "之一的文字";
                }


            }
            if (_targetScoreType == "Comment")
            {
                if (e.FormattedValue.ToString().Length > int.Parse(_commentLimit))
                {
                    cell.ErrorText = "請輸入小於" + _commentLimit + "字數的評語";
                }
            }


        }

        private void dataGridViewX1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                //只檢查分數欄，有錯誤就不給存
                DataGridViewCell cell = dataGridViewX1.Rows[row.Index].Cells[5];

                if (cell.ErrorText != String.Empty)
                {
                    buttonX1.Enabled = false;
                    return; // 有一個錯誤，就不給存，跳出檢查迴圈。
                }
                else
                {
                    buttonX1.Enabled = true;
                }

            }

        }
    }
}
