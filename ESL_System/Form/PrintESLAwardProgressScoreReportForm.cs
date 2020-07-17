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
using Aspose.Cells;
using System.Xml.Linq;
using System.IO;

namespace ESL_System.Form
{
    public partial class PrintESLAwardProgressScoreReportForm : BaseForm
    {
        // 大於多少人數
        private int _moreThan;

        // 大於多少人數 取幾人
        private int _moreThanPeople;

        // 大於多少人數
        private int _lessThan;

        // 小於多少人數 取幾人
        private int _lessThanPeople;

        private List<string> _courseIDList;

        private List<K12.Data.CourseRecord> _courseList;

        private List<ESLCourseRecord> _eslCourseList;

        private List<string> _refAssessmentSetupIDList;

        private BackgroundWorker _worker;

        // 目前學年(106、107、108)
        private string school_year;

        // 目前學期(1、2)
        private string semester;

        // 儲放學生基本資料的 dict ，結構: Ref_studentID,StudentRecord
        private Dictionary<string, ESLAwardStudentRecord> _studentRecordDict = new Dictionary<string, ESLAwardStudentRecord>();

        // 儲放課程 ESL term1 成績的dict ，相同ESL成績評分樣版的課程放在一起整理 其結構為 <assessmentSetupID <courseID,<ESLScore>>> 
        private Dictionary<string, Dictionary<string, List<ESLScore>>> _courseTerm1ScoreDict = new Dictionary<string, Dictionary<string, List<ESLScore>>>();

        // 儲放課程 ESL term2 成績的dict ，相同ESL成績評分樣版的課程放在一起整理 其結構為 <assessmentSetupID <courseID,<ESLScore>>> 
        private Dictionary<string, Dictionary<string, List<ESLScore>>> _courseTerm2ScoreDict = new Dictionary<string, Dictionary<string, List<ESLScore>>>();

        // 儲放課程 ESL 進步成績的dict ，相同ESL成績評分樣版的課程放在一起整理， 其分數 為 term2 與 term 1 的差值 其結構為 <assessmentSetupID <courseID,<ESLScore>>> 
        private Dictionary<string, Dictionary<string, List<ESLScore>>> _courseProgressScoreDict = new Dictionary<string, Dictionary<string, List<ESLScore>>>();

        // 儲放課程 ESL assessment成績的dict ，相同ESL成績評分樣版的課程放在一起整理 其結構為 <assessmentSetupID <courseID,<ESLScore>>> 
        private Dictionary<string, Dictionary<string, List<ESLScore>>> _courseAssessmentScoreDict = new Dictionary<string, Dictionary<string, List<ESLScore>>>();

        // 儲放課程 ESL 總成績的dict ，相同ESL成績評分樣版的課程放在一起整理 其結構為 <assessmentSetupID <courseID,<ESLScore>>> 
        private Dictionary<string, Dictionary<string, List<ESLScore>>> _courseScoreDict = new Dictionary<string, Dictionary<string, List<ESLScore>>>();

        // 儲放課程 ESL 課程修課人數的dict 其結構為 <courseID,<>>
        private Dictionary<string, List<K12.Data.SCAttendRecord>> _courseScattendDict = new Dictionary<string, List<K12.Data.SCAttendRecord>>();

        // 儲放ESL 成績單 科目、比重設定 的dict 其結構為 <courseID,<key,value>>
        private Dictionary<string, Dictionary<string, string>> _itemDict = new Dictionary<string, Dictionary<string, string>>();

        // 儲放每一個評分樣版 對應的 設定結構
        private Dictionary<string, List<Term>> _assessmentSetupDataTableDict = new Dictionary<string, List<Term>>();

        // 評分樣版 id 對應 其樣板名稱
        private Dictionary<string, string> _assessmentSetupIDNamePairDict = new Dictionary<string, string>();

        // 修課課程id 對應的評分樣版 id 對照
        private Dictionary<string, string> _courseIDPairDict = new Dictionary<string, string>();

        // 紀錄成績 為 指標型indicator 的 key值 ， 作為對照 key 為 courseID_termName_subjectName_assessment_Name
        private List<string> _indicatorList = new List<string>();

        // 紀錄成績 為 評語型comment 的 key值
        private List<string> _commentList = new List<string>();

        // 使用者 選擇的 系統試別名稱1
        private string _examName1;

        // 使用者 選擇的 系統試別名稱2
        private string _examName2;

        // 使用者 選擇的 系統試別1 ID
        private string _examID1 = "";

        // 使用者 選擇的 系統試別1 ID
        private string _examID2 = "";


        // 儲放每一個ESL 評分樣版  系統的ExamID 對應到 ESL Term 名稱， 格式 <ESL_TemplateID <ExamID,TermName>>
        // 有了對照才可以 正確在所有ESL 成績中抓到  正確對應目前系統設定的 指定評量成績
        private Dictionary<string, Dictionary<string, string>> _assessmentSetupExamTermDict = new Dictionary<string, Dictionary<string, string>>();

        // 系統Exam 中文名稱 對應 Exam ID
        Dictionary<string, string> ExamDict = new Dictionary<string, string>();


        public PrintESLAwardProgressScoreReportForm(List<string> courseIDList)
        {
            InitializeComponent();
            _courseIDList = courseIDList;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            buttonX1.Enabled = false; // 關閉按鈕

            //驗證 依據 課程修課人數 取的排名人數
            // 是否符合邏輯
            _moreThan = int.Parse("" + numericUpDown1.Value);
            _moreThanPeople = int.Parse("" + numericUpDown2.Value);
            _lessThan = int.Parse("" + numericUpDown3.Value);
            _lessThanPeople = int.Parse("" + numericUpDown4.Value);

            if (_moreThan < _moreThanPeople || _lessThan < _lessThanPeople)
            {
                MsgBox.Show("學生人數下限、學生人數上限，不得少與取得名次數量");
                return;
            };

            if (_moreThan <= _lessThan)
            {
                MsgBox.Show("學生人數下限 不得小於等於學生人數上限");
                return;
            }

            if (comboBoxEx1.Text == comboBoxEx2.Text)
            {
                MsgBox.Show("進步獎項，兩次比較試別不得相同!");
                return;
            }

            // 驗證完畢，開始列印報表
            PrintReport(_courseIDList);
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CheckCalculateTermForm_Load(object sender, EventArgs e)
        {

            // 取得系統中所有的設定評量
            List<K12.Data.ExamRecord> examList = K12.Data.Exam.SelectAll();

            foreach (K12.Data.ExamRecord exam in examList)
            {
                comboBoxEx1.Items.Add(exam.Name);

                comboBoxEx2.Items.Add(exam.Name);

                // 建立Exam對照
                if (!ExamDict.ContainsKey(exam.Name))
                {
                    ExamDict.Add(exam.Name, exam.ID);
                }
            }

            comboBoxEx1.SelectedIndex = 0;
            comboBoxEx2.SelectedIndex = 1;


            string courseIDs = string.Join(",", _courseIDList);

            // 若使用者選取課程沒有 ESL 的樣板設定 則提醒
            string query = @"
SELECT 
    course.ref_exam_template_id
    ,course.id
    ,course.course_name
    ,exam_template.name
    ,exam_template.description 
FROM course 
LEFT JOIN  exam_template ON course.ref_exam_template_id =exam_template.id  
WHERE course.id IN( " + courseIDs + ")";

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query);
            if (dt.Rows.Count > 0)
            {
                List<string> errorList = new List<string>();

                foreach (DataRow dr in dt.Rows)
                {
                    if ("" + dr["description"] == "")
                    {
                        errorList.Add("所選取課程:「" + dr["course_name"] + "」，其並非使用ESL評分樣版，無法使用本功能。 ");
                    }

                    // 將選擇的課程 依照評分樣版 分類 建立 Term1、Term2 結構
                    if (!_courseTerm1ScoreDict.ContainsKey("" + dr["ref_exam_template_id"]))
                    {
                        _courseTerm1ScoreDict.Add("" + dr["ref_exam_template_id"], new Dictionary<string, List<ESLScore>>());

                        _courseTerm1ScoreDict["" + dr["ref_exam_template_id"]].Add("" + dr["id"], new List<ESLScore>());
                    }
                    else
                    {
                        _courseTerm1ScoreDict["" + dr["ref_exam_template_id"]].Add("" + dr["id"], new List<ESLScore>());
                    }

                    if (!_courseTerm2ScoreDict.ContainsKey("" + dr["ref_exam_template_id"]))
                    {
                        _courseTerm2ScoreDict.Add("" + dr["ref_exam_template_id"], new Dictionary<string, List<ESLScore>>());

                        _courseTerm2ScoreDict["" + dr["ref_exam_template_id"]].Add("" + dr["id"], new List<ESLScore>());
                    }
                    else
                    {
                        _courseTerm2ScoreDict["" + dr["ref_exam_template_id"]].Add("" + dr["id"], new List<ESLScore>());
                    }

                    if (!_courseProgressScoreDict.ContainsKey("" + dr["ref_exam_template_id"]))
                    {
                        _courseProgressScoreDict.Add("" + dr["ref_exam_template_id"], new Dictionary<string, List<ESLScore>>());

                        _courseProgressScoreDict["" + dr["ref_exam_template_id"]].Add("" + dr["id"], new List<ESLScore>());
                    }
                    else
                    {
                        _courseProgressScoreDict["" + dr["ref_exam_template_id"]].Add("" + dr["id"], new List<ESLScore>());
                    }

                    // 將選擇的課程 依照評分樣版 分類 建立
                    if (!_courseAssessmentScoreDict.ContainsKey("" + dr["ref_exam_template_id"]))
                    {
                        _courseAssessmentScoreDict.Add("" + dr["ref_exam_template_id"], new Dictionary<string, List<ESLScore>>());

                        _courseAssessmentScoreDict["" + dr["ref_exam_template_id"]].Add("" + dr["id"], new List<ESLScore>());
                    }
                    else
                    {
                        _courseAssessmentScoreDict["" + dr["ref_exam_template_id"]].Add("" + dr["id"], new List<ESLScore>());
                    }
                }

                if (errorList.Count > 0)
                {
                    string erroor = string.Join("\r\n", errorList);

                    MsgBox.Show(erroor);

                    this.Close();
                }
            }
        }


        // 列印 ESL 報表
        private void PrintReport(List<string> courseIDList)
        {
            _examName1 = comboBoxEx1.Text;
            _examName2 = comboBoxEx2.Text;

            _examID1 = ExamDict[_examName1];
            _examID2 = ExamDict[_examName2];

            _worker = new BackgroundWorker();
            _worker.DoWork += new DoWorkEventHandler(Worker_DoWork);
            _worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Worker_RunWorkerCompleted);
            _worker.ProgressChanged += new ProgressChangedEventHandler(Worker_ProgressChanged);
            _worker.WorkerReportsProgress = true;

            _worker.RunWorkerAsync();

        }


        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            _worker.ReportProgress(0, "開始列印 ESL課程進步獎報表...");


            #region 取得目前 系統的 學年度 學期 
            school_year = K12.Data.School.DefaultSchoolYear;
            semester = K12.Data.School.DefaultSemester;
            #endregion


            #region 取得課程 設定樣板、 基本資料整理
            _courseList = new List<K12.Data.CourseRecord>();

            _eslCourseList = new List<ESLCourseRecord>();

            _courseList = K12.Data.Course.SelectByIDs(_courseIDList);

            _eslCourseList = ESLCourseRecord.ToESLCourseRecords(_courseList);

            _refAssessmentSetupIDList = new List<string>();

            foreach (K12.Data.CourseRecord courseRecord in _courseList)
            {
                if (!_refAssessmentSetupIDList.Contains("'" + courseRecord.RefAssessmentSetupID + "'"))
                {
                    _refAssessmentSetupIDList.Add("'" + courseRecord.RefAssessmentSetupID + "'");
                }

                if (!_courseIDPairDict.ContainsKey(courseRecord.ID))
                {
                    _courseIDPairDict.Add(courseRecord.ID, courseRecord.RefAssessmentSetupID);
                }

                if (!_assessmentSetupIDNamePairDict.ContainsKey(courseRecord.RefAssessmentSetupID))
                {
                    _assessmentSetupIDNamePairDict.Add(courseRecord.RefAssessmentSetupID, courseRecord.AssessmentSetup.Name);
                }
            }

            string assessmentSetupIDs = string.Join(",", _refAssessmentSetupIDList);



            #endregion


            #region 建立課程成績 Dict、取得各課程的學生人數(sc_attend)

            List<K12.Data.SCAttendRecord> scList = K12.Data.SCAttend.SelectByCourseIDs(_courseIDList);

            List<string> studentIDList = new List<string>();


            foreach (K12.Data.SCAttendRecord scr in scList)
            {
                studentIDList.Add(scr.Student.ID);

                // 建立課程修課名單，以對照出修課人數 對照畫面設定 決定取幾名 ，
                if (!_courseScattendDict.ContainsKey(scr.Course.ID))
                {
                    _courseScattendDict.Add(scr.Course.ID, new List<K12.Data.SCAttendRecord>());
                    _courseScattendDict[scr.Course.ID].Add(scr);
                }
                else
                {
                    _courseScattendDict[scr.Course.ID].Add(scr);
                }
            }


            #endregion

            // 建立 解讀　description　XML
            CreateFieldTemplate();

            #region 取得、整理ESL成績
            _worker.ReportProgress(20, "取得ESL課程成績");

            List<String> _courseIDListBatch = new List<string>();

            QueryHelper qh = new QueryHelper();
            DataTable dt = new DataTable();

            // 2019/4/09 穎驊優化， 原本的取法，可能會因為筆數過多 造成伺服器資源過載，在此處做分批處理優化(10筆課程一次查詢)
            for (int i = 0; i < _courseIDList.Count; i++)
            {
                if (_courseIDListBatch.Count <= 9 && i + 1 != _courseIDList.Count)
                {
                    _courseIDListBatch.Add(_courseIDList[i]);
                }
                else
                {
                    _courseIDListBatch.Add(_courseIDList[i]);

                    string course_ids = string.Join("','", _courseIDListBatch);

                    string sql = @"
SELECT 
    course.ref_exam_template_id
    ,course.course_name AS english_class
    ,course.id AS course_id 
    ,student.student_number AS student_number
    ,student.name AS student_chinese_name
    ,student.english_name AS student_english_name
    ,student.gender AS gender
    ,class.class_name AS home_room
    ,student.ref_class_id AS ref_class_id
    ,student.id AS student_id
     ,teacher.teacher_name
     ,sc_attend.id AS sc_attend_id
    ,$esl.gradebook_assessment_score.ref_teacher_id
    ,$esl.gradebook_assessment_score.ref_course_id
    ,$esl.gradebook_assessment_score.ref_student_id
    ,$esl.gradebook_assessment_score.term
    ,$esl.gradebook_assessment_score.subject
    ,$esl.gradebook_assessment_score.assessment
    ,$esl.gradebook_assessment_score.custom_assessment
    ,$esl.gradebook_assessment_score.value 
FROM $esl.gradebook_assessment_score  
    LEFT JOIN sc_attend ON $esl.gradebook_assessment_score.ref_sc_attend_id = sc_attend.id 
    LEFT JOIN course ON sc_attend.ref_course_id = course.id
    LEFT JOIN student ON sc_attend.ref_student_id = student.id
    LEFT JOIN class ON student.ref_class_id = class.id
    LEFT JOIN teacher ON $esl.gradebook_assessment_score.ref_teacher_id = teacher.id    
WHERE $esl.gradebook_assessment_score.ref_course_id IN ('" + course_ids + @"')
ORDER BY $esl.gradebook_assessment_score.last_update";


                    _courseIDListBatch.Clear();

                    try
                    {
                        DataTable dt_batch = new DataTable();
                        dt_batch = qh.Select(sql);

                        dt.Merge(dt_batch);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }


            foreach (DataRow row in dt.Rows)
            {
                string termWord = "" + row["term"];
                string subjectWord = "" + row["subject"];
                string assessmentWord = "" + row["assessment"];

                // 課程ID
                string course_id = "" + row["ref_course_id"];

                // 評分樣版 ID
                string assessmentSetupID = "" + row["ref_exam_template_id"];

                string targetTerm1 = _assessmentSetupExamTermDict["" + row["ref_exam_template_id"]].ContainsKey(_examID1) ? _assessmentSetupExamTermDict["" + row["ref_exam_template_id"]][_examID1] : "";

                string targetTerm2 = _assessmentSetupExamTermDict["" + row["ref_exam_template_id"]].ContainsKey(_examID2) ? _assessmentSetupExamTermDict["" + row["ref_exam_template_id"]][_examID2] : "";

                // 有教師自訂的子項目成績就跳掉 不處理
                if ("" + row["custom_assessment"] != "")
                {
                    continue;
                }

                // 當 抓下來的成績 都不是 指定的試別 時 就跳過
                if (targetTerm1 != termWord && targetTerm2 != termWord)
                {
                    continue;
                }

                // 指標型成績
                if (_indicatorList.Contains("" + row["ref_course_id"] + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_')))
                {
                    //獎項只排 分數型成績、指標不納入
                    continue;
                }
                // 評語型成績
                else if (_commentList.Contains("" + row["ref_course_id"] + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_')))
                {
                    //獎項只排 分數型成績、評語不納入
                    continue;
                }

                // 上列狀況排除 剩下的都是分數型成績

                #region 分數整理
                ESLScore eslScore = new ESLScore();

                eslScore.RefStudentID = "" + row["student_id"];
                eslScore.RefStudentName = "" + row["student_chinese_name"];
                eslScore.Term = termWord;
                eslScore.Subject = subjectWord;
                eslScore.Assessment = assessmentWord;

                if ("" + row["value"] != "")
                {
                    eslScore.Score = decimal.Parse("" + row["value"]);
                }

                // 項目都有，為assessment 成績
                if (termWord != "" && "" + subjectWord != "" && "" + assessmentWord != "")
                {
                    _courseAssessmentScoreDict[assessmentSetupID][course_id].Add(eslScore);
                }

                // 沒有assessment，為subject 成績
                if (termWord != "" && "" + subjectWord != "" && "" + assessmentWord == "")
                {
                    // 敘獎 沒有看為subject 成績，暫不處理
                }
                // 沒有assessment、subject，為term 成績1
                if (termWord == targetTerm1 && "" + subjectWord == "" && "" + assessmentWord == "")
                {
                    _courseTerm1ScoreDict[assessmentSetupID][course_id].Add(eslScore);
                }

                // 沒有assessment、subject，為term 成績2
                if (termWord == targetTerm2 && "" + subjectWord == "" && "" + assessmentWord == "")
                {
                    _courseTerm2ScoreDict[assessmentSetupID][course_id].Add(eslScore);
                }
                #endregion

                #region 學生基本資料整理

                if (!_studentRecordDict.ContainsKey("" + row["ref_student_id"]))
                {
                    ESLAwardStudentRecord studentRecord = new ESLAwardStudentRecord();

                    // 學號 (Student Number)
                    studentRecord.StudentNumber = "" + row["student_number"];

                    // 英文姓名 (English Name)
                    studentRecord.EnglishName = "" + row["student_english_name"];

                    // 中文姓名 (Chinese Name)
                    studentRecord.ChineseName = "" + row["student_chinese_name"];

                    // 性別 (Gender)
                    studentRecord.Gender = "" + row["gender"] != "" ? "" + row["gender"] == "1" ? "男" : "女" : "";

                    // 原班級 (Home Room)                
                    studentRecord.HomeRoom = "" + row["home_room"];

                    // 以ref_student_id 當作Key值 加入 對照字典， 之後填資料可以使用
                    _studentRecordDict.Add("" + row["ref_student_id"], studentRecord);
                }

                #endregion

                // 穎驊注解，另外 在樣板中 還有 Level 、 Group ， 目前 2019/1/3 系統中沒有這兩個欄位，
                // 目前預計是等 寒假，在補齊欄位

            }
            #endregion

            _worker.ReportProgress(60, "成績排序中...");

            #region 排序
            // 每一個課程的  Term 成績List 排序
            foreach (string assessmentSetupID in _courseTerm2ScoreDict.Keys)
            {
                foreach (string coursrID in _courseTerm2ScoreDict[assessmentSetupID].Keys)
                {
                    foreach (ESLScore eslscoreTerm2 in _courseTerm2ScoreDict[assessmentSetupID][coursrID])
                    {
                        ESLScore eslscoreTerm1 = _courseTerm1ScoreDict[assessmentSetupID][coursrID].Find(x => x.RefStudentID == eslscoreTerm2.RefStudentID);

                        if (eslscoreTerm1 != null)
                        {
                            // 進步的分數
                            ESLScore progressScore = new ESLScore();

                            progressScore.RefStudentID = eslscoreTerm2.RefStudentID;
                            progressScore.RefStudentName = eslscoreTerm2.RefStudentName;
                            progressScore.Term = eslscoreTerm2.Term;
                            progressScore.Subject = eslscoreTerm2.Subject;
                            progressScore.Assessment = eslscoreTerm2.Assessment;

                            // 進步的分數為 兩次試別 相減
                            progressScore.Score = eslscoreTerm2.Score - eslscoreTerm1.Score;

                            _courseProgressScoreDict[assessmentSetupID][coursrID].Add(progressScore);
                        }
                    }
                }
            }

            // 每一個課程的  Term 成績List 排序
            foreach (string assessmentSetupID in _courseProgressScoreDict.Keys)
            {
                foreach (string coursrID in _courseProgressScoreDict[assessmentSetupID].Keys)
                {
                    // 填 -x 由大排到小 (100、99、98...)
                    try
                    {
                        _courseProgressScoreDict[assessmentSetupID][coursrID].Sort((x, y) => { return -x.Score.CompareTo(y.Score); });
                    }
                    catch
                    {

                    }
                }
            }



            #endregion

            _worker.ReportProgress(80, "填寫報表...");

            // 取得 系統預設的樣板
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.ESL課程班級取前N名_樣板_));

            #region 填表

            FillProgressScoreExcelColunm(wb);

            #endregion

            // 把當作樣板的 第一張 移掉
            wb.Worksheets.RemoveAt(0);

            e.Result = wb;

            _worker.ReportProgress(100, "ESL 報表列印完成。");
        }


        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MsgBox.Show("計算進步失敗!!，錯誤訊息:" + e.Error.Message);
                FISCA.Presentation.MotherForm.SetStatusBarMessage(" ESL課程進步獎產生產生失敗");
                this.Close();
                return;
            }
            else if (e.Cancelled)
            {
                //MsgBox.Show("");
                return;
            }
            else
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage(" ESL課程進步獎產生完成");
            }



            Workbook wb = (Workbook)e.Result;


            // 電子報表功能先暫時不製做
            #region 電子報表
            //// 檢查是否上傳電子報表
            //if (chkUploadEPaper.Checked)
            //{
            //    List<Document> docList = new List<Document>();
            //    foreach (Section ss in doc.Sections)
            //    {
            //        Document dc = new Document();
            //        dc.Sections.Clear();
            //        dc.Sections.Add(dc.ImportNode(ss, true));
            //        docList.Add(dc);
            //    }

            //    Update_ePaper up = new Update_ePaper(docList, "超額比序項目積分證明單", PrefixStudent.系統編號);
            //    if (up.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
            //    {
            //        MsgBox.Show("電子報表已上傳!!");
            //    }
            //    else
            //    {
            //        MsgBox.Show("已取消!!");
            //    }
            //} 
            #endregion

            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "ESL課程進步獎.xlsx";
            sd.Filter = "Excel 檔案(*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    wb.Save(sd.FileName, SaveFormat.Xlsx);
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }

            this.Close();
        }


        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("" + e.UserState, e.ProgressPercentage);
        }



        // 生成 系統Exam 與 ESL Term 成績的對照
        private void CreateFieldTemplate()
        {

            #region  解讀　description　XML

            // 取得ESL 描述 in description
            DataTable dt;
            QueryHelper qh = new QueryHelper();

            string assessmentSetupIDs = string.Join(",", _refAssessmentSetupIDList);

            string selQuery = "SELECT id,description FROM exam_template WHERE id IN (" + assessmentSetupIDs + ")";

            dt = qh.Select(selQuery);

            //整理目前的ESL 課程資料
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    List<Term> termList = new List<Term>();

                    string xmlStr = "<root>" + dr["description"].ToString() + "</root>";
                    XElement elmRoot = XElement.Parse(xmlStr);

                    termList = TransferDescription(elmRoot);

                    foreach (Term term in termList)
                    {
                        if (!_assessmentSetupExamTermDict.ContainsKey(dr["id"].ToString()))
                        {
                            _assessmentSetupExamTermDict.Add(dr["id"].ToString(), new Dictionary<string, string>());

                            _assessmentSetupExamTermDict[dr["id"].ToString()].Add(term.Ref_exam_id, term.Name);
                        }
                        else
                        {
                            if (!_assessmentSetupExamTermDict.ContainsKey(term.Ref_exam_id))
                            {
                                _assessmentSetupExamTermDict[dr["id"].ToString()].Add(term.Ref_exam_id, term.Name);
                            }
                        }
                    }


                    GetAssessmentSetup(termList, _courseList.FindAll(x => x.AssessmentSetup.ID == "" + dr["id"]));

                    if (!_assessmentSetupDataTableDict.ContainsKey("" + dr["id"]))
                    {
                        _assessmentSetupDataTableDict.Add("" + dr["id"], termList);
                    }

                }
            }


            #endregion


        }


        private List<Term> TransferDescription(XElement elmRoot)
        {
            List<Term> tlist = new List<Term>();

            //解析讀下來的 descriptiony 資料，打包成物件群 
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

                        t.Ref_exam_id = ele_term.Attribute("Ref_exam_id") != null ? ele_term.Attribute("Ref_exam_id").Value : ""; // 2018/09/26 穎驊新增，因應要將 ESL 評量導入 成績系統，恩正說，在此加入評量對照

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

                        tlist.Add(t); // 整理成大包的termList 後面用此來拚功能變數總表
                    }
                }
            }


            return tlist;

        }


        private void GetAssessmentSetup(List<Term> termList, List<K12.Data.CourseRecord> couseList)
        {

            foreach (K12.Data.CourseRecord course in couseList)
            {
                foreach (Term term in termList)
                {

                    foreach (Subject subject in term.SubjectList)
                    {
                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            if (assessment.Type == "Score") //分數型成績 才增加
                            {

                            }
                            if (assessment.Type == "Indicator") // 檢查看有沒有　　Indicator　，有的話另外存List 做對照
                            {
                                string key = course.ID + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_');

                                // 作為對照 key 為 courseID_termName_subjectName_assessment_Name
                                if (!_indicatorList.Contains(key))
                                {
                                    _indicatorList.Add(key);
                                }
                            }

                            if (assessment.Type == "Comment") // 檢查看有沒有　　Comment　，有的話另外存List 做對照
                            {
                                string key = course.ID + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_');

                                // 作為對照 key 為 courseID_termName_subjectName_assessment_Name
                                if (!_commentList.Contains(key))
                                {
                                    _commentList.Add(key);
                                }
                            }
                        }
                    }

                }
            }
        }

        // 填寫 課程成績EXCEL
        private void FillProgressScoreExcelColunm(Workbook wb)
        {
            int wscount = 1;

            // 一種ESL樣板 開一個 Worksheet
            foreach (string assessmentSetupID in _courseProgressScoreDict.Keys)
            {
                Worksheet ws = wb.Worksheets[wb.Worksheets.Add()];

                // 複製樣板
                ws.Copy(wb.Worksheets["樣板一"]);

                // 一種樣板 一個sheet 名稱
                string wsName = wscount + ".ESL樣板_" + _assessmentSetupIDNamePairDict[assessmentSetupID] + "_課程";

                wscount++;

                // excel sheet Name 最多只能 31 個字
                wsName = wsName.Length > 26 ? wsName.Substring(0, 22) + "..." : wsName;

                ws.Name = wsName;

                #region 填表頭
                // 填表頭 
                // 標題
                ws.Cells["H1"].Value = school_year + "學年度 第" + semester + "學期 學生成績一覽表";

                // 列印時間
                ws.Cells["L1"].Value = "列印日期:" + DateTime.Now.Date.ToShortDateString();
                #endregion

                // 把範例樣板的成績結構刪除
                ws.Cells.ClearRange(1, 12, 2, 19);

                // 最後 term 的位置
                int termCol = 0;

                // 最後 progressScore 的位置
                int progressScoreCol = 0;

                // 本樣板中有敘獎的學生人數
                int totalAwardsCount = 0;


                int col = 11;

                // 填入每一個 樣板的 成績結構
                foreach (Term term in _assessmentSetupDataTableDict[assessmentSetupID])
                {
                    string targetTerm1 = _assessmentSetupExamTermDict[assessmentSetupID].ContainsKey(_examID1) ? _assessmentSetupExamTermDict[assessmentSetupID][_examID1] : "";

                    string targetTerm2 = _assessmentSetupExamTermDict[assessmentSetupID].ContainsKey(_examID2) ? _assessmentSetupExamTermDict[assessmentSetupID][_examID2] : "";

                    // 不是 使用者選的 兩次評量 項目 跳過
                    if (term.Name != targetTerm1 && term.Name != targetTerm2)
                    {
                        break;
                    }

                    foreach (Subject subject in term.SubjectList)
                    {
                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            //分數型成績
                            if (assessment.Type == "Score")
                            {
                                Cell cell = ws.Cells[1, col];
                                cell.Copy(wb.Worksheets["樣板一"].Cells["M2"]);

                                cell.Value = "(" + subject.Name + ")\n" + assessment.Name;

                                col++;
                            }
                        }
                    }

                    // 最後補上 term
                    Cell cell_term = ws.Cells[1, col];
                    cell_term.Copy(wb.Worksheets["樣板一"].Cells["M2"]);

                    cell_term.Value = term.Name;

                    termCol = col;

                    col++;
                }

                progressScoreCol = termCol + 1;

                // 最後補上 semesterScore
                Cell cell_semesterScore = ws.Cells[1, col];
                cell_semesterScore.Copy(wb.Worksheets["樣板一"].Cells["M2"]);

                cell_semesterScore.Value = "Progress";

                #region 總表樣板
                Worksheet ws_total = wb.Worksheets[wb.Worksheets.Add()];

                // 複製樣板
                ws_total.Copy(ws);

                ws_total.Name = wsName + "(總表)";
                #endregion

                // 依據 表頭的名稱 填入分數
                foreach (string coursrID in _courseProgressScoreDict[assessmentSetupID].Keys)
                {
                    // 取幾名
                    int rankedNumber = 0;

                    // 總人數大於下限 人數
                    if (_courseScattendDict[coursrID].Count >= _moreThan)
                    {
                        rankedNumber = _moreThanPeople;
                    }

                    // 總人數大於上限 人數
                    if (_courseScattendDict[coursrID].Count <= _lessThan)
                    {
                        rankedNumber = _lessThanPeople;
                    }


                    // 依照設定看要取幾名
                    for (int i = 0; i < rankedNumber; i++)
                    {
                        ESLCourseRecord eslCourseRecord = _eslCourseList.First(x => x.ESLID == coursrID);

                        // 學生 系統 ID 
                        string ref_studentID;

                        // 分數
                        if (_courseProgressScoreDict[assessmentSetupID][coursrID].Count >= i + 1)
                        {
                            if (_courseProgressScoreDict[assessmentSetupID][coursrID][i].Score == null)
                            {

                                break;
                            }

                            ref_studentID = _courseProgressScoreDict[assessmentSetupID][coursrID][i].RefStudentID;

                            // 自樣板 把資料第一Row 的格式都Copy
                            ws.Cells.CopyRows(wb.Worksheets["樣板一"].Cells, 2, totalAwardsCount + 2, 1);

                            //進步 分數
                            Cell cell = ws.Cells[totalAwardsCount + 2, progressScoreCol];
                            //cell.Copy(wb.Worksheets["樣板一"].Cells["T2"]);
                            // 先清空值， 預設是沒有分數
                            cell.Value = "N/A";

                            //進步 分數 值
                            cell.Value = _courseProgressScoreDict[assessmentSetupID][coursrID][i].Score;

                            // 如果目前分數 與下一筆分數 同分， 則再增額選進，直到 沒有同分
                            if (_courseProgressScoreDict[assessmentSetupID][coursrID].Count > i + 1)
                            {
                                if (_courseProgressScoreDict[assessmentSetupID][coursrID][i].Score != null && _courseProgressScoreDict[assessmentSetupID][coursrID][i].Score == _courseProgressScoreDict[assessmentSetupID][coursrID][i + 1].Score)
                                {
                                    rankedNumber++;
                                }
                            }

                            // 初始欄位置
                            int initialCol = 11;

                            // 填入每一個 樣板的 成績
                            foreach (Term term in _assessmentSetupDataTableDict[assessmentSetupID])
                            {
                                string targetTerm1 = _assessmentSetupExamTermDict[assessmentSetupID].ContainsKey(_examID1) ? _assessmentSetupExamTermDict[assessmentSetupID][_examID1] : "";

                                string targetTerm2 = _assessmentSetupExamTermDict[assessmentSetupID].ContainsKey(_examID2) ? _assessmentSetupExamTermDict[assessmentSetupID][_examID2] : "";

                                // 不是 使用者選的 兩次評量 項目 跳過
                                if (term.Name != targetTerm1 && term.Name != targetTerm2)
                                {
                                    break;
                                }

                                string termName = term.Name;

                                int assessmentTotal = 0;

                                foreach (Subject subject in term.SubjectList)
                                {
                                    foreach (Assessment assessment in subject.AssessmentList)
                                    {
                                        // 計算一個Term 之下 有幾個 分數 Assessment
                                        if (assessment.Type == "Score")
                                        {
                                            assessmentTotal++;
                                        }

                                    }
                                }

                                // 填 assesssment 成績
                                for (int assesssmentCol = initialCol; assesssmentCol < initialCol + assessmentTotal; assesssmentCol++)
                                {
                                    string subjectAssesssmentName = "" + ws.Cells[1, assesssmentCol].Value;

                                    Cell assesssmentCell = ws.Cells[totalAwardsCount + 2, assesssmentCol];

                                    assesssmentCell.Value = "N/A";

                                    foreach (ESLScore eslScore in _courseAssessmentScoreDict[assessmentSetupID][coursrID])
                                    {
                                        string key = "(" + eslScore.Subject + ")\n" + eslScore.Assessment;
                                        if (eslScore.RefStudentID == ref_studentID && eslScore.Term == termName && key == subjectAssesssmentName)
                                        {
                                            assesssmentCell.Value = eslScore.Score;
                                        }
                                    }
                                }

                                // 填 term 成績
                                Cell termCell = ws.Cells[totalAwardsCount + 2, initialCol + assessmentTotal];

                                termCell.Value = "N/A";

                                foreach (ESLScore eslScore in _courseTerm1ScoreDict[assessmentSetupID][coursrID])
                                {
                                    if (eslScore.RefStudentID == ref_studentID && eslScore.Term == termName)
                                    {
                                        termCell.Value = eslScore.Score;
                                    }
                                }

                                foreach (ESLScore eslScore in _courseTerm2ScoreDict[assessmentSetupID][coursrID])
                                {
                                    if (eslScore.RefStudentID == ref_studentID && eslScore.Term == termName)
                                    {
                                        termCell.Value = eslScore.Score;
                                    }
                                }

                                initialCol = initialCol + assessmentTotal + 1;
                            }


                        }
                        else
                        {
                            // 人數 不足， 跳離迴圈
                            break;
                        }

                        // 學號 (Student Number)
                        ws.Cells[totalAwardsCount + 2, 0].Value = _studentRecordDict[ref_studentID].StudentNumber;

                        // 英文姓名 (English Name)
                        ws.Cells[totalAwardsCount + 2, 1].Value = _studentRecordDict[ref_studentID].EnglishName;

                        // 中文姓名 (Chinese Name)
                        ws.Cells[totalAwardsCount + 2, 2].Value = _studentRecordDict[ref_studentID].ChineseName;

                        // 性別 (Gender)
                        ws.Cells[totalAwardsCount + 2, 3].Value = _studentRecordDict[ref_studentID].Gender;

                        // 原班級 (Home Room)  
                        ws.Cells[totalAwardsCount + 2, 4].Value = _studentRecordDict[ref_studentID].HomeRoom;

                        //課程難度(Level)
                        ws.Cells[totalAwardsCount + 2, 5].Value = eslCourseRecord.ESLDifficulty;

                        // 課程名稱
                        ws.Cells[totalAwardsCount + 2, 7].Value = eslCourseRecord.ESLName;
                        // 教師一
                        ws.Cells[totalAwardsCount + 2, 8].Value = eslCourseRecord.ESLTeachers.Count > 0 ? eslCourseRecord.ESLTeachers.Find(t => t.Sequence == 1).TeacherName : "";
                        // 教師二
                        ws.Cells[totalAwardsCount + 2, 9].Value = eslCourseRecord.ESLTeachers.Count > 1 ? eslCourseRecord.ESLTeachers.Find(t => t.Sequence == 2).TeacherName : "";
                        // 教師三
                        ws.Cells[totalAwardsCount + 2, 10].Value = eslCourseRecord.ESLTeachers.Count > 2 ? eslCourseRecord.ESLTeachers.Find(t => t.Sequence == 3).TeacherName : "";

                        // 穎驊注解，另外 在樣板中 還有 Level 、 Group ， 目前 2019/1/3 系統中沒有這兩個欄位，
                        // 目前預計是等 寒假，在補齊課程欄位


                        totalAwardsCount++;
                    }
                }

                //把多餘的右半邊CELL欄位 砍掉                
                ws.Cells.ClearRange(1, progressScoreCol + 1, totalAwardsCount + 2, 50);
                ws.AutoFitColumns();
                ws.FirstVisibleColumn = 0;// 將打開的介面 調到最左， 要不然就會看到 右邊一片空白。

                totalAwardsCount = 0;

                // 依據 表頭的名稱 填入分數 (總表)
                foreach (string coursrID in _courseProgressScoreDict[assessmentSetupID].Keys)
                {
                    // 取幾名
                    int rankedNumber = 0;

                    rankedNumber = _courseProgressScoreDict[assessmentSetupID][coursrID].Count;

                    // 依照設定看要取幾名
                    for (int i = 0; i < rankedNumber; i++)
                    {
                        ESLCourseRecord eslCourseRecord = _eslCourseList.First(x => x.ESLID == coursrID);

                        // 學生 系統 ID 
                        string ref_studentID;

                        // 分數
                        if (_courseProgressScoreDict[assessmentSetupID][coursrID].Count >= i + 1)
                        {
                            if (_courseProgressScoreDict[assessmentSetupID][coursrID][i].Score == null)
                            {
                                break;
                            }

                            ref_studentID = _courseProgressScoreDict[assessmentSetupID][coursrID][i].RefStudentID;

                            // 自樣板 把資料第一Row 的格式都Copy
                            ws_total.Cells.CopyRows(wb.Worksheets["樣板一"].Cells, 2, totalAwardsCount + 2, 1);

                            //進步 分數
                            Cell cell = ws_total.Cells[totalAwardsCount + 2, progressScoreCol];
                            //cell.Copy(wb.Worksheets["樣板一"].Cells["T2"]);
                            // 先清空值， 預設是沒有分數
                            cell.Value = "N/A";

                            //進步 分數 值
                            cell.Value = _courseProgressScoreDict[assessmentSetupID][coursrID][i].Score;

                            // 如果目前分數 與下一筆分數 同分， 則再增額選進，直到 沒有同分
                            if (_courseProgressScoreDict[assessmentSetupID][coursrID].Count > i + 1)
                            {
                                if (_courseProgressScoreDict[assessmentSetupID][coursrID][i].Score != null && _courseProgressScoreDict[assessmentSetupID][coursrID][i].Score == _courseProgressScoreDict[assessmentSetupID][coursrID][i + 1].Score)
                                {
                                    rankedNumber++;
                                }
                            }

                            // 初始欄位置
                            int initialCol = 11;

                            // 填入每一個 樣板的 成績
                            foreach (Term term in _assessmentSetupDataTableDict[assessmentSetupID])
                            {
                                string targetTerm1 = _assessmentSetupExamTermDict[assessmentSetupID].ContainsKey(_examID1) ? _assessmentSetupExamTermDict[assessmentSetupID][_examID1] : "";

                                string targetTerm2 = _assessmentSetupExamTermDict[assessmentSetupID].ContainsKey(_examID2) ? _assessmentSetupExamTermDict[assessmentSetupID][_examID2] : "";

                                // 不是 使用者選的 兩次評量 項目 跳過
                                if (term.Name != targetTerm1 && term.Name != targetTerm2)
                                {
                                    break;
                                }

                                string termName = term.Name;

                                int assessmentTotal = 0;

                                foreach (Subject subject in term.SubjectList)
                                {
                                    foreach (Assessment assessment in subject.AssessmentList)
                                    {
                                        // 計算一個Term 之下 有幾個 分數 Assessment
                                        if (assessment.Type == "Score")
                                        {
                                            assessmentTotal++;
                                        }

                                    }
                                }

                                // 填 assesssment 成績
                                for (int assesssmentCol = initialCol; assesssmentCol < initialCol + assessmentTotal; assesssmentCol++)
                                {
                                    string subjectAssesssmentName = "" + ws.Cells[1, assesssmentCol].Value;

                                    Cell assesssmentCell = ws_total.Cells[totalAwardsCount + 2, assesssmentCol];

                                    assesssmentCell.Value = "N/A";

                                    foreach (ESLScore eslScore in _courseAssessmentScoreDict[assessmentSetupID][coursrID])
                                    {
                                        string key = "(" + eslScore.Subject + ")\n" + eslScore.Assessment;
                                        if (eslScore.RefStudentID == ref_studentID && eslScore.Term == termName && key == subjectAssesssmentName)
                                        {
                                            assesssmentCell.Value = eslScore.Score;
                                        }
                                    }
                                }

                                // 填 term 成績
                                Cell termCell = ws_total.Cells[totalAwardsCount + 2, initialCol + assessmentTotal];

                                termCell.Value = "N/A";

                                foreach (ESLScore eslScore in _courseTerm1ScoreDict[assessmentSetupID][coursrID])
                                {
                                    if (eslScore.RefStudentID == ref_studentID && eslScore.Term == termName)
                                    {
                                        termCell.Value = eslScore.Score;
                                    }
                                }

                                foreach (ESLScore eslScore in _courseTerm2ScoreDict[assessmentSetupID][coursrID])
                                {
                                    if (eslScore.RefStudentID == ref_studentID && eslScore.Term == termName)
                                    {
                                        termCell.Value = eslScore.Score;
                                    }
                                }

                                initialCol = initialCol + assessmentTotal + 1;
                            }


                        }
                        else
                        {
                            // 人數 不足， 跳離迴圈
                            break;
                        }

                        // 學號 (Student Number)
                        ws_total.Cells[totalAwardsCount + 2, 0].Value = _studentRecordDict[ref_studentID].StudentNumber;

                        // 英文姓名 (English Name)
                        ws_total.Cells[totalAwardsCount + 2, 1].Value = _studentRecordDict[ref_studentID].EnglishName;

                        // 中文姓名 (Chinese Name)
                        ws_total.Cells[totalAwardsCount + 2, 2].Value = _studentRecordDict[ref_studentID].ChineseName;

                        // 性別 (Gender)
                        ws_total.Cells[totalAwardsCount + 2, 3].Value = _studentRecordDict[ref_studentID].Gender;

                        // 原班級 (Home Room)  
                        ws_total.Cells[totalAwardsCount + 2, 4].Value = _studentRecordDict[ref_studentID].HomeRoom;

                        //課程難度(Level)
                        ws_total.Cells[totalAwardsCount + 2, 5].Value = eslCourseRecord.ESLDifficulty;

                        // 課程名稱
                        ws_total.Cells[totalAwardsCount + 2, 7].Value = eslCourseRecord.ESLName;
                        // 教師一
                        ws_total.Cells[totalAwardsCount + 2, 8].Value = eslCourseRecord.ESLTeachers.Count > 0 ? eslCourseRecord.ESLTeachers.Find(t => t.Sequence == 1).TeacherName : "";
                        // 教師二
                        ws_total.Cells[totalAwardsCount + 2, 9].Value = eslCourseRecord.ESLTeachers.Count > 1 ? eslCourseRecord.ESLTeachers.Find(t => t.Sequence == 2).TeacherName : "";
                        // 教師三
                        ws_total.Cells[totalAwardsCount + 2, 10].Value = eslCourseRecord.ESLTeachers.Count > 2 ? eslCourseRecord.ESLTeachers.Find(t => t.Sequence == 3).TeacherName : "";

                        // 穎驊注解，另外 在樣板中 還有 Level 、 Group ， 目前 2019/1/3 系統中沒有這兩個欄位，
                        // 目前預計是等 寒假，在補齊課程欄位


                        totalAwardsCount++;
                    }
                }

                //把多餘的右半邊CELL欄位 砍掉 (總表)             
                ws_total.Cells.ClearRange(1, progressScoreCol + 1, totalAwardsCount + 2, 50);
                ws_total.AutoFitColumns();
                ws_total.FirstVisibleColumn = 0;// 將打開的介面 調到最左， 要不然就會看到 右邊一片空白。
            }

        }

    }



}



