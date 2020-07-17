using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using Aspose.Cells;

using K12.Data;
using System.Xml;
using System.Data;
using FISCA.Data;
using System.Xml.Linq;
using System.ComponentModel;
using FISCA.Presentation.Controls;

namespace ESL_System.Form
{
    public partial class ESLCourseScoreStatusForm : FISCA.Presentation.Controls.BaseForm
    {

        

        private List<CourseDataGridViewRow> _courseDataGridViewRowList;

        private List<string> _CourseIDList;
        private List<ESLCourse> _ESLCourseList = new List<ESLCourse>();

        private BackgroundWorker _worker;

        // 目標樣板ID
        private string _targetTemplateID;

        // 目標試別(Term)名稱
        private string _targetTermName;

        // 目標課程ID (listItem tag 、DataGridViewRow Tag)
        private string _targetCourseID;

        // 目標科目名稱 (DataGridViewColumn tag)
        string _targetSubjectName;

        // 目標評量名稱 (DataGridViewCell tag)
        string _targetAssessmentName; 

        //  ESL 課程ID 與 課程名稱 的對照
        private Dictionary<string, string> _ESLCourseIDNameDict = new Dictionary<string, string>();

        //  ESL 課程ID 與 評分樣版ID 的對照
        private Dictionary<string, string> _ESLCourseIDExamTermIDDict = new Dictionary<string, string>();

        //  評分樣版名稱 與 評分樣版ID 的對照
        private Dictionary<string, string> _ExamTemNameExamTermIDDict = new Dictionary<string, string>();

        //  <評分樣版ID,ESLTemplate>
        private Dictionary<string, ESLTemplate> _ESLTemplateDict = new Dictionary<string, ESLTemplate>();

        //  學生修課資料
        private List<K12.Data.SCAttendRecord> _scaList = new List<SCAttendRecord>();

        // ESL 學生assessment 成績結構 <courseID,<subjectName,List<scoreItem>>>
        private Dictionary<string, Dictionary<string, List<ESLScore>>> _scoreDict = new Dictionary<string, Dictionary<string, List<ESLScore>>>();


        public ESLCourseScoreStatusForm(List<string> eslCouseList)
        {
            InitializeComponent();
            _CourseIDList = eslCouseList;
            
            _worker = new BackgroundWorker();
            _worker.DoWork += new DoWorkEventHandler(Worker_DoWork);
            _worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Worker_RunWorkerCompleted);
            _worker.ProgressChanged += new ProgressChangedEventHandler(Worker_ProgressChanged);
            _worker.WorkerReportsProgress = true;

            GetESLTemplate();

            FillCboTemplate();

            chkDisplayNotFinish.Enabled = false;
            picLoading.Visible = false;

        }

        private void GetESLTemplate()
        {
            picLoading.Visible = true;

            string courseIDs = string.Join(",", _CourseIDList);

            #region 取得ESL 課程資料
            // 2018/06/12 抓取課程且其有ESL 樣板設定規則的，才做後續整理，  在table exam_template 欄位 description 不為空代表其為ESL 的樣板
            string query = @"
                    SELECT 
                        course.id AS courseID
                        ,course.course_name
                        ,exam_template.description 
                        ,exam_template.id AS templateID
                        ,exam_template.name AS templateName
                    FROM course 
                    LEFT JOIN  exam_template ON course.ref_exam_template_id =exam_template.id  
                    WHERE course.id IN( " + courseIDs + ") AND  exam_template.description IS NOT NULL  ";

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


            #region 取得學生修課資料

            _scaList = K12.Data.SCAttend.SelectByCourseIDs(_CourseIDList);

            #endregion

        }


        private void FillCboTemplate()
        {
            picLoading.Visible = true;

            cboTemplate.Items.Clear();

            foreach (string templateID in _ESLTemplateDict.Keys)
            {
                cboTemplate.Items.Add(_ESLTemplateDict[templateID].ESLTemplateName);

                _ExamTemNameExamTermIDDict.Add(_ESLTemplateDict[templateID].ESLTemplateName, _ESLTemplateDict[templateID].ID);
            }
        }


        /// <summary>
        /// 當試別改變時觸發
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
           
            #region DataGridView  新介面
            // 清內容
            dataGridViewX1.Rows.Clear();

            // 清表頭
            dataGridViewX1.Columns.Clear();

            // 課程名稱表頭
            DataGridViewColumn colCourseName = new DataGridViewColumn();

            colCourseName.CellTemplate = new DataGridViewTextBoxCell();
            colCourseName.Name = "colCourseName";
            colCourseName.HeaderText = "課程名稱";
            colCourseName.ReadOnly = true;
            colCourseName.Width = 200;
            dataGridViewX1.Columns.Add(colCourseName);

            // 填寫完畢項目表頭
            DataGridViewColumn colTotalStatus = new DataGridViewColumn();

            colTotalStatus.CellTemplate = new DataGridViewTextBoxCell();
            colTotalStatus.Name = "colTotalStatus";
            colTotalStatus.HeaderText = "填寫完畢項目";
            colTotalStatus.ReadOnly = true;
            colTotalStatus.Width = 130;
            dataGridViewX1.Columns.Add(colTotalStatus);


            // 依目前 所選樣板、試別 動態產生 DataGridView 表頭
            foreach (Term term in _ESLTemplateDict[_targetTemplateID].TermList)
            {
                if (term.Name == _targetTermName)
                {
                    foreach (Subject subject in term.SubjectList)
                    {
                        foreach (Assessment assessment in subject.AssessmentList)
                        {

                            DataGridViewColumn col = new DataGridViewColumn();

                            col.CellTemplate = new DataGridViewTextBoxCell();
                            col.Name = "Col_" + subject.Name + "_" + assessment.Name;
                            col.HeaderText = subject.Name + " / " + assessment.Name;
                            col.ReadOnly = true;
                            col.Width = (subject.Name.Length + assessment.Name.Length) * 9;

                            col.Tag = subject.Name; // 將subject 資訊 tag 放在col

                            dataGridViewX1.Columns.Add(col);
                        }
                    }
                }
            } 
            #endregion


            // 暫停畫面控制項
            chkDisplayNotFinish.Enabled = false;
            cboTemplate.SuspendLayout();
            cboExam.SuspendLayout();
            dataGridViewX1.SuspendLayout();

            _worker.RunWorkerAsync();
        }


        /// <summary>
        /// 依試別取得所有課程成績
        /// </summary>
        /// <param name="targetTermName"></param>
        private void LoadCourses(string targetTermName)
        {

            _worker.ReportProgress(0, "取得課程資料...");

            List<string> targetCourseList = new List<string>();

            // 在所有選擇課程中 其樣板 為目前選擇樣板 才加入查詢成績
            foreach (string courseID in _CourseIDList)
            {
                if (_ESLCourseIDExamTermIDDict[courseID] == _targetTemplateID)
                {
                    targetCourseList.Add(courseID);
                }
            }

            string targetCourseIDs = string.Join(",", targetCourseList);



            #region 建立應有成績名單



            _scoreDict.Clear(); // 每次重抓都把成績Dict 清乾淨

            int scaCount = 0;

            // 以之前抓到的學生修課名單，建立應該要有的成績資料
            foreach (K12.Data.SCAttendRecord scaRecord in _scaList)
            {
                // 假如本筆修課紀錄 的課程 非設定為本次篩選ESL樣板 跳過
                if (_ESLCourseIDExamTermIDDict[scaRecord.RefCourseID] != _targetTemplateID)
                {
                    continue;
                }

                // 2018/11/14 穎驊修正，若學生有修課紀錄， 但是目前 狀態 為非一般，則不顯示。
                if (scaRecord.Student.Status != StudentRecord.StudentStatus.一般)
                {
                    continue;
                }


                _worker.ReportProgress(10 + 50 * (scaCount++ / _scaList.Count), "取得修課學生資料...");

                // 目標樣板設定
                List<Term> termList = _ESLTemplateDict[_targetTemplateID].TermList;

                if (!_scoreDict.ContainsKey(scaRecord.RefCourseID))
                {
                    _scoreDict.Add(scaRecord.RefCourseID, new Dictionary<string, List<ESLScore>>());

                    foreach (Term term in termList)
                    {
                        if (term.Name == _targetTermName)
                        {
                            foreach (Subject subject in term.SubjectList)
                            {
                                foreach (Assessment assessment in subject.AssessmentList)
                                {

                                    // 取得授課教師
                                    CourseTeacherRecord teacher = scaRecord.Course.Teachers.Find(t => t.Sequence == int.Parse(assessment.TeacherSequence));

                                    ESLScore scoreItem = new ESLScore();

                                    scoreItem.Term = term.Name;
                                    scoreItem.Subject = subject.Name;
                                    scoreItem.Assessment = assessment.Name;

                                    scoreItem.RefCourseID = scaRecord.RefCourseID;
                                    scoreItem.RefStudentID = scaRecord.RefStudentID;
                                    scoreItem.RefScAttendID = scaRecord.ID;  // 參考修課紀錄ID(依據ESL2019寒假優化，成績參考修課紀錄ID，將廢除RefCourseID、RefStudentID)
                                    scoreItem.RefTeacherID = teacher != null ? teacher.TeacherID : ""; // 教師ID

                                    scoreItem.RefCourseName = scaRecord.Course.Name;                                    
                                    scoreItem.RefTeacherName = teacher != null ? teacher.TeacherName : ""; ; // 教師名稱
                                    scoreItem.RefStudentName = scaRecord.Student.Name;

                                    scoreItem.HasValue = false; // 一開始都先當 教師沒有輸入成績，等到取得成績後 再回填scoreItem

                                    if (!_scoreDict[scaRecord.RefCourseID].ContainsKey(subject.Name))
                                    {
                                        _scoreDict[scaRecord.RefCourseID].Add(subject.Name, new List<ESLScore>());
                                        _scoreDict[scaRecord.RefCourseID][subject.Name].Add(scoreItem);
                                    }
                                    else
                                    {
                                        _scoreDict[scaRecord.RefCourseID][subject.Name].Add(scoreItem);
                                    }

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
                                    CourseTeacherRecord teacher = scaRecord.Course.Teachers.Find(t => t.Sequence == int.Parse(assessment.TeacherSequence));

                                    ESLScore scoreItem = new ESLScore();

                                    scoreItem.Term = term.Name;
                                    scoreItem.Subject = subject.Name;
                                    scoreItem.Assessment = assessment.Name;

                                    scoreItem.RefCourseID = scaRecord.RefCourseID;
                                    scoreItem.RefStudentID = scaRecord.RefStudentID;
                                    scoreItem.RefScAttendID = scaRecord.ID;  // 參考修課紀錄ID(依據ESL2019寒假優化，成績參考修課紀錄ID，將廢除RefCourseID、RefStudentID)
                                    scoreItem.RefTeacherID = teacher != null ? teacher.TeacherID : ""; // 教師ID

                                    scoreItem.RefCourseName = scaRecord.Course.Name;
                                    scoreItem.RefTeacherName = teacher != null ? teacher.TeacherName : ""; ; // 教師名稱
                                    scoreItem.RefStudentName = scaRecord.Student.Name;

                                    scoreItem.HasValue = false; // 一開始都先當 教師沒有輸入成績，等到取得成績後 再回填scoreItem

                                    if (!_scoreDict[scaRecord.RefCourseID].ContainsKey(subject.Name))
                                    {
                                        _scoreDict[scaRecord.RefCourseID].Add(subject.Name, new List<ESLScore>());
                                        _scoreDict[scaRecord.RefCourseID][subject.Name].Add(scoreItem);
                                    }
                                    else
                                    {
                                        _scoreDict[scaRecord.RefCourseID][subject.Name].Add(scoreItem);
                                    }

                                }
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
                    AND sc_attend.ref_course_id IN( " + targetCourseIDs + ")";

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query);



            //整理目前的ESL 課程資料

            _worker.ReportProgress(60, "取得學生成績資料...");

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
                            foreach (string subjectName in _scoreDict[courseID].Keys)
                            {
                                if (subjectName == "" + dr["subject"])
                                {
                                    foreach (ESLScore scoreItem in _scoreDict[courseID][subjectName])
                                    {
                                        if (scoreItem.Assessment == "" + dr["assessment"] && scoreItem.RefStudentID == "" + dr["ref_student_id"] )
                                        {
                                            scoreItem.ID = "" + dr["uid"]; // 填入 uid 之後可以做為更新使用

                                            scoreItem.Value = "" + dr["value"]; // 填分數

                                            scoreItem.HasValue = true; // 若以上條件可找到配對，則本課程、本term、subject、assessment 的教師 有輸入成績。

                                        }
                                    }
                                }
                            }
                        }
                    }

                }
            }
            #endregion

            int scoreCount = 0;
            
            // 填 DataGridView
            _courseDataGridViewRowList = new List< CourseDataGridViewRow>();

            foreach (string courseID in _scoreDict.Keys)
            {

                _worker.ReportProgress(60 + 30 * (scoreCount++ / _scoreDict.Keys.Count), "取得學生成績資料...");

                // 假如該課程 為採用目前所選 樣板
                if (_ESLCourseIDExamTermIDDict[courseID] == _targetTemplateID)
                {
                    CourseDataGridViewRow row = new CourseDataGridViewRow(_ESLCourseIDNameDict[courseID], _scoreDict[courseID], _ESLTemplateDict[_targetTemplateID], _targetTermName);

                    row.Tag = courseID; // 用課程ID 當作 Tag

                    _courseDataGridViewRowList.Add(row);
                }
            }
        }



        /// <summary>
        /// 將課程填入 DataGridView
        /// </summary>
        private void FillCourses(List<CourseDataGridViewRow> list)
        {
            if (list.Count <= 0) return;


            dataGridViewX1.SuspendLayout();
            dataGridViewX1.Rows.Clear();
            dataGridViewX1.Rows.AddRange(list.ToArray());
            dataGridViewX1.ResumeLayout();

        }



        /// <summary>
        /// 按下「關閉」時觸發
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 改變「僅顯示未完成輸入之課程」時觸發
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkDisplayNotFinish_CheckedChanged(object sender, EventArgs e)
        {            
            FillCourses(GetDisplayDataGridViewList());
        }




        /// <summary>
        /// 取得要顯示的 CourseListViewItemList
        /// </summary>
        /// <returns></returns>
        private List<CourseDataGridViewRow> GetDisplayDataGridViewList()
        {
            if (chkDisplayNotFinish.Checked == true)
            {
                List<CourseDataGridViewRow> list = new List<CourseDataGridViewRow>();
                foreach (CourseDataGridViewRow item in _courseDataGridViewRowList)
                {
                    if (item.IsFinish) continue;
                    list.Add(item);
                }
                return list;
            }
            else
            {
                return _courseDataGridViewRowList;
            }
        }

        /// <summary>
        /// 按下「匯出到 Excel」時觸發
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.Worksheets.Clear();
            Worksheet ws = book.Worksheets[book.Worksheets.Add()];
            ws.Name = "ESL課程成績輸入檢視.xls";

            int index = 0;
            Dictionary<string, int> map = new Dictionary<string, int>();

            #region 建立標題
            for (int i = 0; i < dataGridViewX1.Columns.Count; i++)
            {
                DataGridViewColumn col = dataGridViewX1.Columns[i];
                ws.Cells[index, i].PutValue(col.HeaderText);
                map.Add(col.HeaderText, i);
            }
            index++;
            #endregion

            #region 填入內容
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow) continue;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    int column = map[cell.OwningColumn.HeaderText];
                    ws.Cells[index, column].PutValue("" + cell.Value);
                }
                index++;
            }
            #endregion

            SaveFileDialog sd = new SaveFileDialog();
            sd.FileName = "ESL課程成績輸入檢視";
            sd.Filter = "Excel檔案(*.xlsx)|*.xlsx";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                DialogResult result = new DialogResult();

                try
                {
                    book.Save(sd.FileName, SaveFormat.Xlsx);
                    result = MsgBox.Show("檔案儲存完成，是否開啟檔案?", "是否開啟", MessageBoxButtons.YesNo);
                }
                catch (Exception ex)
                {
                    MsgBox.Show("儲存失敗。" + ex.Message);
                }

                if (result == DialogResult.Yes)
                {
                    try
                    {
                        System.Diagnostics.Process.Start(sd.FileName);
                    }
                    catch (Exception ex)
                    {
                        MsgBox.Show("開啟檔案發生失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// 按下「重新整理」時觸發
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRefresh_Click(object sender, EventArgs e)
        {

            RefreshListView();
        }



        /// <summary>
        /// 每一筆課程的評量狀況 (DataGridView)
        /// </summary>
        private class CourseDataGridViewRow : DataGridViewRow
        {
            //  已輸入人數 / 總人數 (教師名稱)
            private const string Format = "{0}/{1} ({2})";

            private bool _is_finish;
            public bool IsFinish { get { return _is_finish; } }


            // 紀錄每一個 subject_assessment 的 分子
            private Dictionary<string, int> assessmentCountDict = new Dictionary<string, int>();

            // 紀錄每一個 subject_assessment 的 total 分母
            private Dictionary<string, int> assessmentTotalCountDict = new Dictionary<string, int>();

            // 紀錄每一個 subject_assessment 與老師名字的配對
            private Dictionary<string, string> assessmentTeacherNametDict = new Dictionary<string, string>();


            //每一次 傳一筆 Course 有的 subjectDict 進來，還有目前 template ，targetTermName
            public CourseDataGridViewRow(string courseName, Dictionary<string, List<ESLScore>> subjectDict, ESLTemplate ESLTemplateDict, string targetTermName)
            {
                _is_finish = true;

                // 數出 個項目的 輸入情況
                foreach (Term term in ESLTemplateDict.TermList)
                {
                    if (term.Name == targetTermName)
                    {
                        foreach (Subject subject in term.SubjectList)
                        {
                            foreach (Assessment assessment in subject.AssessmentList)
                            {
                                foreach (string subjectName in subjectDict.Keys)
                                {
                                    if (subject.Name == subjectName)
                                    {
                                        foreach (ESLScore scoreItem in subjectDict[subjectName])
                                        {
                                            if (assessment.Name == scoreItem.Assessment)
                                            {
                                                //有值 才加分子 ，且不能為空字串(web 前端成績輸入會存到如此資料，恩正說，此狀況視為沒有輸入)
                                                if (scoreItem.HasValue && scoreItem.Value != "")
                                                {
                                                    // 子項目 分子
                                                    if (!assessmentCountDict.ContainsKey(subject.Name + "_" + assessment.Name))
                                                    {
                                                        assessmentCountDict.Add(subject.Name + "_" + assessment.Name, 1);
                                                    }
                                                    else
                                                    {
                                                        assessmentCountDict[subject.Name + "_" + assessment.Name]++;
                                                    }
                                                }
                                                else
                                                {
                                                    // 子項目 分子
                                                    if (!assessmentCountDict.ContainsKey(subject.Name + "_" + assessment.Name))
                                                    {
                                                        assessmentCountDict.Add(subject.Name + "_" + assessment.Name, 0);
                                                    }

                                                }

                                                // 子項目 分母
                                                if (!assessmentTotalCountDict.ContainsKey(subject.Name + "_" + assessment.Name))
                                                {
                                                    assessmentTotalCountDict.Add(subject.Name + "_" + assessment.Name, 1);
                                                }
                                                else
                                                {
                                                    assessmentTotalCountDict[subject.Name + "_" + assessment.Name]++;
                                                }

                                                // 子項目 老師名稱
                                                if (!assessmentTeacherNametDict.ContainsKey(subject.Name + "_" + assessment.Name))
                                                {
                                                    assessmentTeacherNametDict.Add(subject.Name + "_" + assessment.Name, scoreItem.RefTeacherName);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // 依目前 所選樣板、試別 對出 與表頭 相應的內容

                DataGridViewCell courseNameCell = new DataGridViewTextBoxCell();
                courseNameCell.Value = courseName;
                this.Cells.Add(courseNameCell);

                // 數出總項目 未填寫完畢項目
                foreach (Term term in ESLTemplateDict.TermList)
                {
                    if (term.Name == targetTermName)
                    {
                        int total = 0;
                        int scoreCount = 0;


                        foreach (Subject subject in term.SubjectList)
                        {
                            foreach (Assessment assessment in subject.AssessmentList)
                            {
                                total++;

                                int assessmentTotal = assessmentTotalCountDict[subject.Name + "_" + assessment.Name];
                                int assessmentScoreCount = assessmentCountDict[subject.Name + "_" + assessment.Name];

                                // 子項目 總數 等於 有分數 的數量，就是 完成
                                if (assessmentTotal == assessmentScoreCount)
                                {
                                    scoreCount++;
                                }
                            }
                        }

                        string ScoreField = string.Format("{0}/{1}", scoreCount, total);

                        // 填總項目 已完成 數量

                        DataGridViewCell cell = new DataGridViewTextBoxCell();
                        cell.Value = ScoreField;

                        DataGridViewCellStyle style = new DataGridViewCellStyle();
                        style.ForeColor = (total == scoreCount) ? Color.Black : Color.Red;
                        cell.Style = style;

                        this.Cells.Add(cell);
                    }
                }

                // 子項目 每一個學生分數的填寫 狀態
                foreach (Term term in ESLTemplateDict.TermList)
                {
                    if (term.Name == targetTermName)
                    {
                        foreach (Subject subject in term.SubjectList)
                        {
                            foreach (Assessment assessment in subject.AssessmentList)
                            {
                                int total = assessmentTotalCountDict[subject.Name + "_" + assessment.Name];
                                int scoreCount = assessmentCountDict[subject.Name + "_" + assessment.Name];
                                string ScoreField = string.Format(Format, scoreCount, total, assessmentTeacherNametDict[subject.Name + "_" + assessment.Name]);

                                if (total != scoreCount)
                                {
                                    _is_finish = false;
                                }

                                DataGridViewCell cell = new DataGridViewTextBoxCell();
                                cell.Value = ScoreField;

                                DataGridViewCellStyle style = new DataGridViewCellStyle();
                                style.ForeColor = (total == scoreCount) ? Color.Black : Color.Red;
                                cell.Style = style;
                                cell.Tag = assessment.Name;

                                this.Cells.Add(cell);
                            }
                        }
                    }
                }                
            }


        }

        private void ESLCourseScoreStatusForm_Load(object sender, EventArgs e)
        {

        }

        private void cboTemplate_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            cboExam.Enabled = true;

            _targetTemplateID = _ExamTemNameExamTermIDDict["" + cboTemplate.SelectedItem];

            FillcboExam();
        }

        private void FillcboExam()
        {
            cboExam.Items.Clear();

            foreach (Term term in (_ESLTemplateDict[_targetTemplateID].TermList))
            {
                cboExam.Items.Add(term.Name);
            }
        }


        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {

            LoadCourses(_targetTermName);

        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 繼續 畫面控制項       
            picLoading.Visible = false;
            chkDisplayNotFinish.Enabled = true;            
            cboTemplate.ResumeLayout();
            cboExam.ResumeLayout();
            
            dataGridViewX1.ResumeLayout();
                        
            FillCourses(GetDisplayDataGridViewList());

            FISCA.Presentation.MotherForm.SetStatusBarMessage("取得ESL課程教師輸入狀態完成");
            
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage(""+e.UserState, e.ProgressPercentage);
        }


        private void dataGridViewX1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            if (e.ColumnIndex < 0) return;
            if (e.RowIndex < 0) return;
            DataGridViewCell cell = dataGridViewX1.Rows[e.RowIndex].Cells[e.ColumnIndex];

            _targetCourseID = "" + cell.OwningRow.Tag; //  targetTermName
            _targetSubjectName = "" + cell.OwningColumn.Tag; //  targetSubjectName
            _targetAssessmentName = "" + cell.Tag; //  targetAssessmentName

            if (_targetSubjectName == "" || _targetAssessmentName == "") return;

            Form.ESLScoreInputForm inputForm = new ESLScoreInputForm(_ESLCourseIDNameDict[_targetCourseID], _ESLTemplateDict[_targetTemplateID], _scoreDict[_targetCourseID], _targetTermName, _targetSubjectName, _targetAssessmentName, _scaList);

            inputForm.ShowDialog();

            RefreshListView(); // 更改完成績後，重整畫面
        }
    }
}

