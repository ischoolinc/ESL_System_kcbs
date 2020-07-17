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

using System.Xml;

using FISCA.Data;
using System.Xml.Linq;




namespace ESL_System.Form
{
    public partial class ESLTransferScoreInputForm : FISCA.Presentation.Controls.BaseForm
    {
        // 目標修課紀錄ID
        private string _targetScAttendID;

        // 目標課程ID
        private string _targetCourseID;

        // 目標課程名稱
        private string _targetCourseName;

        // 目標課程樣板ID
        private string _targetTemplateID;

        // 目標試別(Term)名稱
        private string _targetTermName;

        // 指標性評語 可以使用項目
        private List<Indicators> _targetIndicatorList = new List<Indicators>();

        private List<string> _CourseIDList = new List<string>();

        // 目前本課程的 ESL 評分樣版
        private ESLTemplate _eslTemplate;

        // ESL 分數 <subjectName,List<ESL分數>>
        private Dictionary<string, List<ESLScore>> _scoreDict = new Dictionary<string, List<ESLScore>>();

        private BackgroundWorker _uploadWorker; // 上傳成績使用

        private BackgroundWorker _downloadWorker; // 下載成績使用

        //  ESL 課程ID 與 課程名稱 的對照
        private Dictionary<string, string> _ESLCourseIDNameDict = new Dictionary<string, string>();

        //  ESL 課程ID 與 評分樣版ID 的對照
        private Dictionary<string, string> _ESLCourseIDExamTermIDDict = new Dictionary<string, string>();

        //  評分樣版名稱 與 評分樣版ID 的對照
        private Dictionary<string, string> _ExamTemNameExamTermIDDict = new Dictionary<string, string>();

        //  <評分樣版ID,ESLTemplate>
        private Dictionary<string, ESLTemplate> _ESLTemplateDict = new Dictionary<string, ESLTemplate>();

        //  該轉學學生修課資料
        private SCAttendRecord _scAttendRecord = new SCAttendRecord();


        public ESLTransferScoreInputForm(string targetScAttendID)
        {
            InitializeComponent();

            _targetScAttendID = targetScAttendID;

            #region 取得學生修課資料

            _scAttendRecord = K12.Data.SCAttend.SelectByID(_targetScAttendID);

            #endregion

            #region 取得ESL 課程資料            

            _targetCourseID = _scAttendRecord.RefCourseID;

            _targetCourseName = _scAttendRecord.Course.Name;

            _targetTemplateID = _scAttendRecord.Course.RefAssessmentSetupID;

            string query = @"
                    SELECT 
                        course.id AS courseID
                        ,course.course_name
                        ,exam_template.description 
                        ,exam_template.id AS templateID
                        ,exam_template.name AS templateName
                    FROM course 
                    LEFT JOIN  exam_template ON course.ref_exam_template_id =exam_template.id  
                    WHERE course.id IN( " + _targetCourseID + ") AND  exam_template.description IS NOT NULL  ";

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query);

            _CourseIDList.Clear(); // 清空

            //整理目前的ESL 課程資料
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    ESLCourse record = new ESLCourse();

                    _CourseIDList.Add("" + dr["courseID"]); // 加入真正的 是ESL 課程ID

                    if (!_ESLCourseIDExamTermIDDict.ContainsKey("" + dr["courseID"]))
                    {
                        _ESLCourseIDExamTermIDDict.Add("" + dr["courseID"], "" + dr["templateID"]);
                    }

                    if (!_ESLTemplateDict.ContainsKey("" + dr["templateID"]))
                    {
                        ESLTemplate template = new ESLTemplate();

                        template.ID = "" + dr["templateID"];
                        template.ESLTemplateName = "" + dr["templateName"];
                        template.Description = "" + dr["description"];

                        _ESLTemplateDict.Add("" + dr["templateID"], template);
                    }


                    // 建立課程名稱對照
                    if (!_ESLCourseIDNameDict.ContainsKey("" + dr["courseID"]))
                    {
                        _ESLCourseIDNameDict.Add("" + dr["courseID"], "" + dr["course_name"]);
                    }


                }
            }
            #endregion

            #region 解析ESL 課程 計算規則
            // 解析計算規則
            foreach (string templateID in _ESLTemplateDict.Keys)
            {
                string xmlStr = "<root>" + _ESLTemplateDict[templateID].Description + "</root>";
                XElement elmRoot = XElement.Parse(xmlStr);

                //解析讀下來的 descriptiony 資料
                if (elmRoot != null)
                {
                    if (elmRoot.Element("ESLTemplate") != null)
                    {
                        foreach (XElement ele_term in elmRoot.Element("ESLTemplate").Elements("Term"))
                        {
                            Term t = new Term();

                            t.Name = ele_term.Attribute("Name").Value;
                            t.Weight = ele_term.Attribute("Weight").Value;
                            t.InputStartTime = ele_term.Attribute("InputStartTime").Value;
                            t.InputEndTime = ele_term.Attribute("InputEndTime").Value;
                            t.Ref_exam_id = ele_term.Attribute("Ref_exam_id").Value;

                            t.SubjectList = new List<Subject>();

                            foreach (XElement ele_subject in ele_term.Elements("Subject"))
                            {
                                Subject s = new Subject();

                                s.Name = ele_subject.Attribute("Name").Value;
                                s.Weight = ele_subject.Attribute("Weight").Value;

                                s.AssessmentList = new List<Assessment>();

                                foreach (XElement ele_assessment in ele_subject.Elements("Assessment"))
                                {
                                    Assessment a = new Assessment();

                                    a.Name = ele_assessment.Attribute("Name").Value;
                                    a.Weight = ele_assessment.Attribute("Weight").Value;
                                    a.TeacherSequence = ele_assessment.Attribute("TeacherSequence").Value;
                                    a.Type = ele_assessment.Attribute("Type").Value;
                                    a.AllowCustomAssessment = ele_assessment.Attribute("AllowCustomAssessment").Value;

                                    if (a.Type == "Comment") // 假如是 評語類別，多讀一項 輸入限制屬性
                                    {
                                        a.InputLimit = ele_assessment.Attribute("InputLimit").Value;
                                    }

                                    a.IndicatorsList = new List<Indicators>();

                                    if (ele_assessment.Element("Indicators") != null)
                                    {

                                        foreach (XElement ele_Indicator in ele_assessment.Element("Indicators").Elements("Indicator"))
                                        {
                                            Indicators i = new Indicators();

                                            i.Name = ele_Indicator.Attribute("Name").Value;
                                            i.Description = ele_Indicator.Attribute("Description").Value;

                                            a.IndicatorsList.Add(i);
                                        }
                                    }
                                    s.AssessmentList.Add(a);
                                }
                                t.SubjectList.Add(s);
                            }



                            _ESLTemplateDict[templateID].TermList.Add(t);
                        }
                    }
                }
            }
            #endregion


            _eslTemplate = _ESLTemplateDict[_ESLCourseIDExamTermIDDict[_targetCourseID]];

            //_scoreDict = scoreDict;

            labelX1.Text = _targetCourseName;

            // 填 試別
            FillcboExam();

            // 填入分數
            //FillScore();
        }


        private void FillcboExam()
        {
            cboExam.Items.Clear();

            foreach (Term term in (_ESLTemplateDict[_targetTemplateID].TermList))
            {
                cboExam.Items.Add(term.Name);
            }

            cboExam.Enabled = true;
        }

        private void FillScore()
        {
            dataGridViewX1.Rows.Clear();

            foreach (ESLScore scoreItem in _scoreDict[_targetCourseID])
            {
                DataGridViewRow row = new DataGridViewRow();

                row.CreateCells(dataGridViewX1);

                row.Cells[0].Value = _scAttendRecord.Student.Class != null ? _scAttendRecord.Student.Class.Name : ""; // 學生班級
                row.Cells[1].Value = _scAttendRecord.Student != null ? "" + _scAttendRecord.Student.SeatNo : "";  // 學生座號
                row.Cells[2].Value = _scAttendRecord.Student != null ? "" + _scAttendRecord.Student.Name : "";      // 學生姓名
                row.Cells[3].Value = _scAttendRecord.Student != null ? "" + _scAttendRecord.Student.StudentNumber : "";  // 學生學號
                row.Cells[4].Value = scoreItem.Term; //試別
                row.Cells[5].Value = scoreItem.Subject; //科目
                row.Cells[6].Value = scoreItem.Assessment; //評量
                row.Cells[7].Value = scoreItem.RefTeacherName;
                row.Cells[8].Value = scoreItem.Value;
                row.Cells[9].Value = scoreItem.Ratio;

                row.Tag = _scAttendRecord.ID;  // row tag 用sc_attend_id 就夠(依據2019 ESL 寒假優化項目)

                dataGridViewX1.Rows.Add(row);
            }

            picLoading.Visible = false;
        }


        // 儲存
        private void buttonX1_Click(object sender, EventArgs e)
        {
            // 2019/03/12 轉學生 補輸入成績 的驗證機制要再想一下，
            // 一次輸入所有類型(分數、評語、指標)的成績 有點麻煩

            //// 輸入分數 超過 界線提醒
            //bool outRangeWarning = false;

            ////if (_targetScoreType == "Score")
            ////{
            ////    decimal i = 0;

            ////    foreach (DataGridViewRow row in dataGridViewX1.Rows)
            ////    {
            ////        //只檢查分數欄， 其值 若 超出 0~100 的範圍 跳出提醒視窗 是否繼續儲存
            ////        DataGridViewCell cell = dataGridViewX1.Rows[row.Index].Cells[5];

            ////        if (decimal.TryParse("" + cell.Value, out i))
            ////        {
            ////            if (i < 0 || i > 100)
            ////            {
            ////                outRangeWarning = true;
            ////            }
            ////        }
            ////    }
            ////}

            //// 若畫面上有分數超過範圍，則跳出提醒視窗， 讓使用者決定是否要繼續儲存。
            //if (outRangeWarning)
            //{
            //    if (MsgBox.Show("表格上有分數成績不在0~100範圍內，請問是否繼續儲存?", "警告!", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
            //    {
            //        return;
            //    }
            //}

            _uploadWorker = new BackgroundWorker();
            _uploadWorker.DoWork += new DoWorkEventHandler(UploadWorker_DoWork);
            _uploadWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(UploadWorker_RunWorkerCompleted);
            _uploadWorker.ProgressChanged += new ProgressChangedEventHandler(UploadWorker_ProgressChanged);
            _uploadWorker.WorkerReportsProgress = true;

            // 暫停畫面控制項
            dataGridViewX1.SuspendLayout();
            buttonX1.Enabled = false;

            _uploadWorker.RunWorkerAsync();
        }



        private void UploadWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _uploadWorker.ReportProgress(30, "整理成績資料...");

            //拚SQL
            // 兜資料
            List<string> dataList = new List<string>();

            List<ESLScore> updateESLscoreList = new List<ESLScore>(); // 最後要update ESLscoreList
            List<ESLScore> insertESLscoreList = new List<ESLScore>(); // 最後要indert ESLscoreList

            int i = 0;

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                string targetSubjectName = "" + row.Cells[5].Value;
                string targetAssessmentName = "" + row.Cells[6].Value;

                foreach (ESLScore scoreItem in _scoreDict[_targetCourseID])
                {
                    // 科目、評量相同 、 學生ID 相同 則為該成績
                    if (scoreItem.Subject == targetSubjectName && scoreItem.Assessment == targetAssessmentName && scoreItem.RefScAttendID == "" + row.Tag)
                    {
                        string ratio = "" + row.Cells[9].Value;

                        ratio = ratio.Trim().Replace("'", "''"); // trim 掉空白、 單引號特殊字

                        if ("" + ratio != "")
                        {
                            scoreItem.Ratio = int.Parse(ratio); // 填比例
                        }
                        
                        // 本來有成績， 但值不同了 加入 更新名單
                        if (scoreItem.HasValue && scoreItem.Value != "" + row.Cells[8].Value)
                        {
                            scoreItem.OldValue = "" + scoreItem.Value; // 舊分數

                            scoreItem.Value = "" + row.Cells[8].Value; // 新分數

                            scoreItem.Value = scoreItem.Value.Trim().Replace("'", "''"); // trim 掉空白、 單引號特殊字
                            
                            updateESLscoreList.Add(scoreItem);
                        }

                        // 本來沒成績， 本次新填的值 ，加入新增名單
                        if (!scoreItem.HasValue && "" + row.Cells[8].Value != "")
                        {
                            scoreItem.Value = "" + row.Cells[8].Value; // 新分數

                            scoreItem.Value = scoreItem.Value.Trim().Replace("'", "''"); // trim 掉空白、 單引號特殊字

                            scoreItem.HasValue = true;

                            insertESLscoreList.Add(scoreItem);
                        }
                    }
                }

                _uploadWorker.ReportProgress(30 + (i++ * 60 / dataGridViewX1.Rows.Count), "整理成績資料...");

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
                    ,'{10}'::BIGINT AS ratio
                    ,'{11}'::INTEGER AS uid
                    ,'UPDATE'::TEXT AS action
                ", score.RefStudentID, score.RefCourseID, score.RefScAttendID, score.RefTeacherID, score.Term, score.Subject, score.Assessment, "", score.Value, score.OldValue, score.Ratio, score.ID);

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
                    ,'{10}'::BIGINT AS ratio
                    ,'{11}'::INTEGER AS uid
                    ,'INSERT'::TEXT AS action
                ", score.RefStudentID, score.RefCourseID, score.RefScAttendID, score.RefTeacherID, score.Term, score.Subject, score.Assessment, "", score.Value, "", score.Ratio, 0);  // insert 給 oldValue ""、 uid = 0

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
    ,score_data_row.ratio
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
        ,ratio = score_data_row.ratio
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
    ,ratio
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
    ,score_data_row.ratio::BIGINT AS ratio	
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
	, 'ESL後端成績輸入(轉學生)' AS action
	, 'student'::TEXT AS target_category
	, score_data_row_with_log_data.ref_student_id AS target_id
	, now() AS server_time
	, '{2}' AS client_info
	, 'ESL後端成績輸入(轉學生)'AS action_by   
	, '學號「'|| score_data_row_with_log_data.student_number||'」，學生「'|| score_data_row_with_log_data.student_name ||'」，ESL課程「'|| score_data_row_with_log_data.course_name || '」，試別「'||score_data_row_with_log_data.term|| '」，科目 「'||score_data_row_with_log_data.subject|| '」，評量「'||score_data_row_with_log_data.assessment|| '」，新增成績「'||score_data_row_with_log_data.value|| '」 ，人工指定比例「'||score_data_row_with_log_data.ratio|| '」  ' AS description 
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
	, 'ESL後端成績輸入(轉學生)' AS action
	, 'student'::TEXT AS target_category
	, score_data_row_with_log_data.ref_student_id AS target_id
	, now() AS server_time
	, '{2}' AS client_info
	, 'ESL後端成績輸入(轉學生)'AS action_by   
	, '學號「'|| score_data_row_with_log_data.student_number||'」，學生「'|| score_data_row_with_log_data.student_name ||'」，ESL課程「'|| score_data_row_with_log_data.course_name || '」，試別「'||score_data_row_with_log_data.term|| '」，科目 「'||score_data_row_with_log_data.subject|| '」，評量「'||score_data_row_with_log_data.assessment|| '」，成績自「'||score_data_row_with_log_data.old_value|| '」修改為「'||score_data_row_with_log_data.value|| '」，人工指定比例「'||score_data_row_with_log_data.ratio|| '」  ' AS description 
FROM
	score_data_row_with_log_data
WHERE action ='UPDATE'
", Data, _actor, _client_info);



            UpdateHelper uh = new UpdateHelper();

            _uploadWorker.ReportProgress(90, "上傳成績...");

            //執行sql
            try
            {
                uh.Execute(sql);
                _uploadWorker.ReportProgress(100, "ESL 評量成績上傳完成。");
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        private void UploadWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
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

        private void UploadWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("" + e.UserState, e.ProgressPercentage);
        }

        private void dataGridViewX1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            // 只驗第九格 分數、指標、評語 欄位 、第十格 比重
            if (e.ColumnIndex != 8 || e.ColumnIndex != 9)
            {
                return;
            }

            DataGridViewCell cell = dataGridViewX1.Rows[e.RowIndex].Cells[e.ColumnIndex];

            cell.ErrorText = String.Empty;

            //if (_targetScoreType == "Score")
            //{
            //    decimal i = 0;

            //    if (!decimal.TryParse("" + e.FormattedValue, out i))
            //    {
            //        cell.ErrorText = "請輸入數值。";                    
            //    }
            //}
            //if (_targetScoreType == "Indicator")
            //{
            //    if (_targetIndicatorList.Find(indicator => indicator.Name == "" + e.FormattedValue) ==null)
            //    {
            //        List<string> indicatorList = new List<string>();

            //        string indicators = "";

            //        foreach (Indicators indicator in _targetIndicatorList)
            //        {
            //            indicatorList.Add(indicator.Name);
            //        }

            //        indicators = string.Join("、", indicatorList);

            //        cell.ErrorText = "請輸入" + indicators + "之一的文字";
            //    }


            //}
            //if (_targetScoreType == "Comment")
            //{
            //    if (e.FormattedValue.ToString().Length > int.Parse(_commentLimit))
            //    {
            //        cell.ErrorText = "請輸入小於" + _commentLimit + "字數的評語";
            //    }
            //}


        }

        private void dataGridViewX1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            //foreach (DataGridViewRow row in dataGridViewX1.Rows)
            //{
            //    //只檢查分數欄，有錯誤就不給存
            //    DataGridViewCell cell = dataGridViewX1.Rows[row.Index].Cells[5];

            //    if (cell.ErrorText != String.Empty)
            //    {
            //        buttonX1.Enabled = false;
            //        return; // 有一個錯誤，就不給存，跳出檢查迴圈。
            //    }
            //    else
            //    {
            //        buttonX1.Enabled = true;
            //    }

            //}
        }

        private void cboExam_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshListView();
        }

        /// <summary>
        /// 更新ListView
        /// </summary>
        private void RefreshListView()
        {
            picLoading.Visible = true;

            if (cboExam.SelectedItem == null) return;

            _targetTermName = "" + cboExam.SelectedItem;

            _downloadWorker = new BackgroundWorker();
            _downloadWorker.DoWork += new DoWorkEventHandler(DownloadWorker_DoWork);
            _downloadWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(DownloadWorker_RunWorkerCompleted);
            _downloadWorker.ProgressChanged += new ProgressChangedEventHandler(DownloadWorker_ProgressChanged);
            _downloadWorker.WorkerReportsProgress = true;

            _downloadWorker.RunWorkerAsync();
        }


        private void DownloadWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            LoadCourses(_targetTermName);
        }

        private void DownloadWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FillScore();
        }

        private void DownloadWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }



        /// <summary>
        /// 依試別取得所有課程成績
        /// </summary>
        /// <param name="targetTermName"></param>
        private void LoadCourses(string targetTermName)
        {
            _downloadWorker.ReportProgress(0, "取得課程資料...");

            #region 建立該學生應有成績

            _scoreDict.Clear(); // 每次重抓都把成績Dict 清乾淨

            // 目標樣板設定
            List<Term> termList = _ESLTemplateDict[_targetTemplateID].TermList;

            if (!_scoreDict.ContainsKey(_scAttendRecord.RefCourseID))
            {
                _scoreDict.Add(_scAttendRecord.RefCourseID, new List<ESLScore>());

                foreach (Term term in termList)
                {
                    if (term.Name == _targetTermName)
                    {
                        foreach (Subject subject in term.SubjectList)
                        {
                            foreach (Assessment assessment in subject.AssessmentList)
                            {
                                // 取得授課教師
                                CourseTeacherRecord teacher = _scAttendRecord.Course.Teachers.Find(t => t.Sequence == int.Parse(assessment.TeacherSequence));

                                ESLScore scoreItem = new ESLScore();

                                scoreItem.Term = term.Name;
                                scoreItem.Subject = subject.Name;
                                scoreItem.Assessment = assessment.Name;

                                scoreItem.RefCourseID = _scAttendRecord.RefCourseID;
                                scoreItem.RefStudentID = _scAttendRecord.RefStudentID;
                                scoreItem.RefScAttendID = _scAttendRecord.ID;  // 參考修課紀錄ID(依據ESL2019寒假優化，成績參考修課紀錄ID，將廢除RefCourseID、RefStudentID)
                                scoreItem.RefTeacherID = teacher != null ? teacher.TeacherID : ""; // 教師ID

                                scoreItem.RefCourseName = _scAttendRecord.Course.Name;
                                scoreItem.RefTeacherName = teacher != null ? teacher.TeacherName : ""; ; // 教師名稱
                                scoreItem.RefStudentName = _scAttendRecord.Student.Name;

                                scoreItem.HasValue = false; // 一開始都先當 教師沒有輸入成績，等到取得成績後 再回填scoreItem

                                _scoreDict[_scAttendRecord.RefCourseID].Add(scoreItem);
                            }
                        }
                    }
                }
            }
            else
            {
                foreach (Term term in termList)
                {
                    if (term.Name == _targetTermName)
                    {
                        foreach (Subject subject in term.SubjectList)
                        {
                            foreach (Assessment assessment in subject.AssessmentList)
                            {
                                // 取得授課教師
                                CourseTeacherRecord teacher = _scAttendRecord.Course.Teachers.Find(t => t.Sequence == int.Parse(assessment.TeacherSequence));

                                ESLScore scoreItem = new ESLScore();

                                scoreItem.Term = term.Name;
                                scoreItem.Subject = subject.Name;
                                scoreItem.Assessment = assessment.Name;

                                scoreItem.RefCourseID = _scAttendRecord.RefCourseID;
                                scoreItem.RefStudentID = _scAttendRecord.RefStudentID;
                                scoreItem.RefScAttendID = _scAttendRecord.ID;  // 參考修課紀錄ID(依據ESL2019寒假優化，成績參考修課紀錄ID，將廢除RefCourseID、RefStudentID)
                                scoreItem.RefTeacherID = teacher != null ? teacher.TeacherID : ""; // 教師ID

                                scoreItem.RefCourseName = _scAttendRecord.Course.Name;
                                scoreItem.RefTeacherName = teacher != null ? teacher.TeacherName : ""; ; // 教師名稱
                                scoreItem.RefStudentName = _scAttendRecord.Student.Name;

                                scoreItem.HasValue = false; // 一開始都先當 教師沒有輸入成績，等到取得成績後 再回填scoreItem

                                _scoreDict[_scAttendRecord.RefCourseID].Add(scoreItem);
                            }
                        }
                    }
                }
            }

            #endregion


            #region 取得 本試別 ESL 課程成績資料 (參考修課紀錄ID(依據ESL2019寒假優化，成績參考修課紀錄ID，將廢除RefCourseID、RefStudentID))

            string query = @"
                    SELECT
                        $esl.gradebook_assessment_score.uid
                        ,$esl.gradebook_assessment_score.term
                        ,$esl.gradebook_assessment_score.subject
                        ,$esl.gradebook_assessment_score.assessment
                        ,$esl.gradebook_assessment_score.custom_assessment
                        ,$esl.gradebook_assessment_score.value                        
                        ,$esl.gradebook_assessment_score.ratio         
                        ,$esl.gradebook_assessment_score.ref_sc_attend_id
                        ,$esl.gradebook_assessment_score.ref_teacher_id
                        ,course.id AS ref_course_id
                        ,student.id AS ref_student_id 
                        ,course.course_name 
                        ,teacher.teacher_name 
                        ,student.name AS student_name
                    FROM $esl.gradebook_assessment_score 
                    LEFT JOIN  sc_attend ON  $esl.gradebook_assessment_score.ref_sc_attend_id = sc_attend.id  
                    LEFT JOIN  course ON  sc_attend.ref_course_id = course.id                      
                    LEFT JOIN  student ON  sc_attend.ref_student_id= student.id
                    LEFT JOIN  teacher ON  $esl.gradebook_assessment_score.ref_teacher_id= teacher.id  
                    WHERE 
                    assessment IS NOT NULL
                    AND term = '" + targetTermName + @"'
                    AND sc_attend.ref_course_id IN( " + _targetCourseID + ")";

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query);



            //整理目前的ESL 課程資料

            _downloadWorker.ReportProgress(60, "取得學生成績資料...");

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    // 濾掉有 custom_assessment 項目的成績，不用SQL AND custom_assessment!='' 的原因是因為有的時候custom_assessment 會NULL
                    if ("" + dr["custom_assessment"] != "")
                    {
                        continue;
                    }

                    foreach (string courseID in _scoreDict.Keys)
                    {
                        if (courseID == "" + dr["ref_course_id"])
                        {
                            foreach (ESLScore scoreItem in _scoreDict[courseID])
                            {
                                if (scoreItem.Subject == "" + dr["subject"] && scoreItem.Assessment == "" + dr["assessment"] && scoreItem.RefStudentID == "" + dr["ref_student_id"])
                                {
                                    scoreItem.ID = "" + dr["uid"]; // 填入 uid 之後可以做為更新使用

                                    scoreItem.Value = "" + dr["value"]; // 填分數

                                    if ("" + dr["ratio"] != "")
                                    {
                                        scoreItem.Ratio = int.Parse("" + dr["ratio"]); // 填比例
                                    }                                    
                                    scoreItem.HasValue = true; // 若以上條件可找到配對，則本課程、本term、subject、assessment 的教師 有輸入成績。
                                }
                            }
                        }
                    }

                }
            }
            #endregion

        }




    }
}
