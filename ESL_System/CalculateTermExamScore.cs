using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Data;
using DevComponents.DotNetBar;
using System.Xml.Linq;
using K12.Data;
using System.Xml;
using System.Data;
using FISCA.Presentation.Controls;
using System.ComponentModel;

namespace ESL_System
{
    // 2018/10/03 穎驊修正，與恩正討論後，依照data flow diagram 重新設計程式
    class CalculateTermExamScore
    {
        private string target_exam_id; //目標試別id

        // 不再使用， 因應新的調整， 小數位數 將會 跟著每一個樣板設定
        //private int _decimalPlace = 2; // 結算成績 計算的小數精度位數， 基本上為2， 如果後續有調整需要再另外設計架構

        private List<string> _courseIDList;
        private List<ESLCourse> _ESLCourseList = new List<ESLCourse>();

        private BackgroundWorker _worker;

        private Dictionary<string, List<Term>> _scoreTemplateDict = new Dictionary<string, List<Term>>(); // 各課程的分數計算規則 

        private Dictionary<string, int> _scoreDecimalPlaceDict = new Dictionary<string, int>(); // 各課程的計算小數位數 

        private Dictionary<string, decimal> _scoreRatioDict = new Dictionary<string, decimal>(); // 各課程的分數比例權重分子

        private Dictionary<string, string> _scoreExamScoreTypeDict = new Dictionary<string, string>(); // 各課程的分數計算評量成績分數種類(定期、平時)

        private Dictionary<string, decimal> _scoreRatioTotalDict = new Dictionary<string, decimal>(); // 各課程的分數比例權重分母

        private Dictionary<string, ESLScore> _subjectScoreDict = new Dictionary<string, ESLScore>(); // 計算用的科目成績字典
        private Dictionary<string, ESLScore> _termScoreDict = new Dictionary<string, ESLScore>(); // 計算用的評量成績字典

        private List<ESLScore> _eslscoreList = new List<ESLScore>(); // 先暫時這樣儲存 上傳用，之後會想改用scoreUpsertDict

        private List<K12.Data.SCAttendRecord> _scatList; // ESL學生 修課紀錄 List

        private Dictionary<string, List<K12.Data.SCAttendRecord>> _scatDict = new Dictionary<string, List<SCAttendRecord>>(); // ESL 學生修課紀錄 Dict <studentID,List<SCAttendRecord>>

        private Dictionary<string, List<K12.Data.SCETakeRecord>> _scetDict = new Dictionary<string, List<SCETakeRecord>>(); // 學生系統Exam紀錄 Dict <studentID,List<SCETakeRecord>>

        private Dictionary<string, List<SCETakeESLRecord>> _scetESLDict = new Dictionary<string, List<SCETakeESLRecord>>(); // ESL 學生系統Exam紀錄 Dict <studentID,List<SCETakeESLRecord>>

        private Dictionary<string, List<ESLScore>> _scoreAssessmentOriDict = new Dictionary<string, List<ESLScore>>();  // 取得ESL成績ori (assessment)
        private Dictionary<string, List<ESLScore>> _scoreTermSubjectOriDict = new Dictionary<string, List<ESLScore>>();  // 取得ESL成績ori (term、subject)

        private Dictionary<string, List<ESLScore>> _scorefinalDict = new Dictionary<string, List<ESLScore>>();  // 存放計算完的ESL成績final

        private Dictionary<string, List<ESLScore>> _scoreUpdateDict = new Dictionary<string, List<ESLScore>>();  // 存放計算完要update的ESL成績
        private Dictionary<string, List<ESLScore>> _scoreInsertDict = new Dictionary<string, List<ESLScore>>();  // 存放計算完要insert的ESL成績final



        public CalculateTermExamScore(List<string> courseIDList, string exam_id)
        {
            _courseIDList = courseIDList;
            target_exam_id = exam_id;
        }

        public void CalculateESLTermScore()
        {
            _worker = new BackgroundWorker();
            _worker.DoWork += new DoWorkEventHandler(Worker_DoWork);
            _worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Worker_RunWorkerCompleted);
            _worker.ProgressChanged += new ProgressChangedEventHandler(Worker_ProgressChanged);
            _worker.WorkerReportsProgress = true;
            _worker.WorkerSupportsCancellation = true;
            _worker.RunWorkerAsync();

        }



        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {

            _worker.ReportProgress(0, "取得課程資料...");

            List<string> _TargetCourseTermList = new List<string>(); //本次須計算的目標課程Term <TermName>

            string courseIDs = string.Join(",", _courseIDList);

            #region 取得ESL 課程資料
            // 2018/06/12 抓取課程且其有ESL 樣板設定規則的，才做後續整理，  在table exam_template 欄位 description 不為空代表其為ESL 的樣板
            string query = @"
                    SELECT 
                        course.id
                        ,course.course_name
                        ,exam_template.description 
                        ,exam_template.extension
                    FROM course 
                    LEFT JOIN  exam_template ON course.ref_exam_template_id =exam_template.id  
                    WHERE course.id IN( " + courseIDs + ") AND  exam_template.description IS NOT NULL  ";

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query);

            _courseIDList.Clear(); // 清空

            //整理目前的ESL 課程資料
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    ESLCourse record = new ESLCourse();

                    _courseIDList.Add("" + dr[0]); // 加入真正的 是ESL 課程ID

                    record.ID = "" + dr[0]; //課程ID
                    record.CourseName = "" + dr[1]; //課程名稱
                    record.Description = "" + dr[2]; // ESL 評分樣版設定
                    record.Extension = "" + dr[3]; // ESL 評分樣版定期平時占比

                    _ESLCourseList.Add(record);
                }
            }
            #endregion

            #region 取得課程(老師資料)

            List<K12.Data.CourseRecord> courseList = K12.Data.Course.SelectByIDs(_courseIDList);



            #endregion

            _worker.ReportProgress(10, "取得解析ESL課程樣板...");

            #region 解析ESL 課程 計算規則
            // 解析計算規則
            foreach (ESLCourse course in _ESLCourseList)
            {
                int _decimalPlace = 2; // 預設兩位數
                string xmlStr = "<root>" + course.Description + "</root>";
                XElement elmRoot = XElement.Parse(xmlStr);

                //解析讀下來的 descriptiony 資料
                if (elmRoot != null)
                {
                    if (elmRoot.Element("ESLTemplate") != null)
                    {
                        _decimalPlace = elmRoot.Element("ESLTemplate").Attribute("decimalPlace") != null ? int.Parse(elmRoot.Element("ESLTemplate").Attribute("decimalPlace").Value) : 2; // 小數位數預設為2

                        foreach (XElement ele_term in elmRoot.Element("ESLTemplate").Elements("Term"))
                        {
                            Term t = new Term();

                            t.Name = ele_term.Attribute("Name").Value;
                            t.Weight = ele_term.Attribute("Weight").Value;
                            t.InputStartTime = ele_term.Attribute("InputStartTime").Value;
                            t.InputEndTime = ele_term.Attribute("InputEndTime").Value;
                            t.Ref_exam_id = ele_term.Attribute("Ref_exam_id").Value;

                            if (target_exam_id == t.Ref_exam_id && !_TargetCourseTermList.Contains("'" + t.Name + "'"))
                            {
                                _TargetCourseTermList.Add("'" + t.Name + "'");
                            }

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

                                    if (a.Type == "Score") // 假如是 分數類別，多讀一項 評量計算類別 (定期、平時) (若沒有設定，預設為定期)
                                    {
                                        a.ExamScoreType = ele_assessment.Attribute("ExamScoreType") != null ? ele_assessment.Attribute("ExamScoreType").Value : "定期";
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

                            if (!_scoreTemplateDict.ContainsKey(course.ID))
                            {
                                _scoreTemplateDict.Add(course.ID, new List<Term>());

                                _scoreTemplateDict[course.ID].Add(t);
                            }
                            else
                            {
                                _scoreTemplateDict[course.ID].Add(t);
                            }

                            if (!_scoreDecimalPlaceDict.ContainsKey(course.ID))
                            {
                                _scoreDecimalPlaceDict.Add(course.ID, _decimalPlace);
                            }


                        }
                    }
                }

                // 解析 Extension 比例資料
                xmlStr = "<root>" + course.Extension + "</root>";
                elmRoot = XElement.Parse(xmlStr);

                if (elmRoot != null)
                {
                    if (elmRoot.Element("Extension") != null)
                    {
                        if (elmRoot.Element("Extension").Element("ScorePercentage") != null)
                        {
                            _scoreRatioTotalDict.Add(course.ID + "_中文系統評分樣版定期比例", int.Parse(elmRoot.Element("Extension").Element("ScorePercentage").Value));
                            _scoreRatioTotalDict.Add(course.ID + "_中文系統評分樣版平時比例", 100 - int.Parse(elmRoot.Element("Extension").Element("ScorePercentage").Value));
                        }
                    }
                }
            }
            #endregion

            _worker.ReportProgress(20, "取得ESL課程修課學生...");

            #region 取得ESL 課程 修課學生 修課紀錄

            // K12 API 會抓到 畢業及離校 的學生 修課紀錄
            _scatList = K12.Data.SCAttend.SelectByCourseIDs(_courseIDList);

            // 將全部狀態非一般生的學生 修課紀錄移除
            _scatList.RemoveAll(scaRecord => scaRecord.Student.Status != StudentRecord.StudentStatus.一般);


            // 將修課紀錄 以stidentID 整理成字典
            foreach (K12.Data.SCAttendRecord scattendRecord in _scatList)
            {
                if (!_scatDict.ContainsKey(scattendRecord.RefStudentID))
                {
                    _scatDict.Add(scattendRecord.RefStudentID, new List<SCAttendRecord>());

                    _scatDict[scattendRecord.RefStudentID].Add(scattendRecord);
                }
                else
                {
                    _scatDict[scattendRecord.RefStudentID].Add(scattendRecord);
                }
            }

            #endregion

            _worker.ReportProgress(30, "取得ESL課程修課學生成績...");

            #region 取得學生ESL 成績(assessment)
            // 學生ID清單
            //string studentIDs = string.Join(",", _scatDict.Keys);

            // 修課紀錄 sc_attend ID清單
            string sc_attend_IDs = string.Join(",", _scatList.Select(scaRecord => scaRecord.ID).ToList());

            // Term 名稱 清單
            string termNames = string.Join(",", _TargetCourseTermList);

            //抓取目前所選取ESL  課程、評量，且其修課學生的 成績
            //assessment 欄位 為空，還有濾掉 custom_assessment 成績 代表此成績 是 Subject 或是 Term 成績            
            //query = @"SELECT * 
            //          FROM $esl.gradebook_assessment_score 
            //          WHERE ref_course_id IN( " + courseIDs + ") " +
            //          "AND  ref_student_id IN(" + studentIDs + ")" +
            //          "AND term IN(" + termNames + ")" +
            //          "AND assessment != ''";

            // 2019/01/29 ESL 寒假改版， 採用 sc_attend 找資料
            query = @"SELECT 
                      $esl.gradebook_assessment_score.uid
                      ,sc_attend.ref_course_id
                     ,sc_attend.ref_student_id
                     ,$esl.gradebook_assessment_score.ref_teacher_id
                     ,$esl.gradebook_assessment_score.ref_sc_attend_id
                     ,$esl.gradebook_assessment_score.term
                     ,$esl.gradebook_assessment_score.subject
                     ,$esl.gradebook_assessment_score.assessment
                     ,$esl.gradebook_assessment_score.custom_assessment
                     ,$esl.gradebook_assessment_score.value
                     ,$esl.gradebook_assessment_score.ratio
                      FROM $esl.gradebook_assessment_score 
                      LEFT JOIN sc_attend ON sc_attend.id = $esl.gradebook_assessment_score.ref_sc_attend_id
                      WHERE ref_sc_attend_id IN( " + sc_attend_IDs + ") " +
                      "AND term IN(" + termNames + ")" +
                      "AND assessment != ''";

            dt = qh.Select(query);

            // 整理 既有的成績資料 (Assessment)
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    // 濾掉有 custom_assessment 項目的成績，不用SQL AND custom_assessment!='' 的原因是因為有的時候custom_assessment 會NULL
                    if ("" + dr["custom_assessment"] != "")
                    {
                        continue;
                    }

                    if (!_scoreAssessmentOriDict.ContainsKey("" + dr["ref_student_id"]))
                    {
                        ESLScore score = new ESLScore();

                        score.ID = "" + dr["uid"];
                        score.RefCourseID = "" + dr["ref_course_id"];
                        score.RefStudentID = "" + dr["ref_student_id"];
                        score.RefTeacherID = "" + dr["ref_teacher_id"];
                        score.RefScAttendID = "" + dr["ref_sc_attend_id"];
                        score.Term = "" + dr["term"];
                        score.Subject = "" + dr["subject"];
                        score.Assessment = "" + dr["assessment"];
                        score.Custom_Assessment = "" + dr["custom_assessment"];
                        score.Value = "" + dr["value"];

                        // 2019/02/01 穎驊新增，若學生assessment成績，有指定的比例設定，將會保存此屬性，作為權重使用
                        if ("" + dr["ratio"] != "")
                        {
                            score.Ratio = int.Parse("" + dr["ratio"]);
                        }

                        _scoreAssessmentOriDict.Add("" + dr["ref_student_id"], new List<ESLScore>());

                        _scoreAssessmentOriDict["" + dr["ref_student_id"]].Add(score);
                    }
                    else
                    {
                        ESLScore score = new ESLScore();

                        score.ID = "" + dr["uid"];
                        score.RefCourseID = "" + dr["ref_course_id"];
                        score.RefStudentID = "" + dr["ref_student_id"];
                        score.RefTeacherID = "" + dr["ref_teacher_id"];
                        score.RefScAttendID = "" + dr["ref_sc_attend_id"];
                        score.Term = "" + dr["term"];
                        score.Subject = "" + dr["subject"];
                        score.Assessment = "" + dr["assessment"];
                        score.Custom_Assessment = "" + dr["custom_assessment"];
                        score.Value = "" + dr["value"];

                        // 2019/02/01 穎驊新增，若學生assessment成績，有指定的比例設定，將會保存此屬性，作為權重使用
                        if ("" + dr["ratio"] != "")
                        {
                            score.Ratio = int.Parse("" + dr["ratio"]);
                        }

                        _scoreAssessmentOriDict["" + dr["ref_student_id"]].Add(score);
                    }
                }
            }
            else
            {
                e.Cancel = true;
                return;
            }
            #endregion



            #region 取得學生 ESL 成績(term、subject) 作為最後對照是否更新使用

            // 2019/01/29 ESL 寒假改版， 採用 sc_attend 找資料
            query = @"SELECT 
                      $esl.gradebook_assessment_score.uid
                     ,sc_attend.ref_course_id
                     ,sc_attend.ref_student_id
                     ,$esl.gradebook_assessment_score.ref_teacher_id
                     ,$esl.gradebook_assessment_score.ref_sc_attend_id
                     ,$esl.gradebook_assessment_score.term
                     ,$esl.gradebook_assessment_score.subject
                     ,$esl.gradebook_assessment_score.assessment
                     ,$esl.gradebook_assessment_score.custom_assessment
                     ,$esl.gradebook_assessment_score.value
                     ,$esl.gradebook_assessment_score.ratio
                      FROM $esl.gradebook_assessment_score 
                      LEFT JOIN sc_attend ON sc_attend.id = $esl.gradebook_assessment_score.ref_sc_attend_id" +
                      " WHERE ref_sc_attend_id IN( " + sc_attend_IDs + ") " +
                      " AND term IN(" + termNames + ")" +
                      " AND assessment IS NULL ";

            dt = qh.Select(query);

            // 整理 既有的成績資料 
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    // 濾掉有 custom_assessment 項目的成績，不用SQL AND custom_assessment!='' 的原因是因為有的時候custom_assessment 會NULL
                    if ("" + dr["custom_assessment"] != "")
                    {
                        continue;
                    }

                    if (!_scoreTermSubjectOriDict.ContainsKey("" + dr["ref_student_id"]))
                    {
                        ESLScore score = new ESLScore();

                        score.ID = "" + dr["uid"];
                        score.RefCourseID = "" + dr["ref_course_id"];
                        score.RefStudentID = "" + dr["ref_student_id"];
                        score.RefTeacherID = "" + dr["ref_teacher_id"];
                        score.RefScAttendID = "" + dr["ref_sc_attend_id"];
                        score.Term = "" + dr["term"];
                        score.Subject = "" + dr["subject"];
                        score.Assessment = "" + dr["assessment"];
                        score.Custom_Assessment = "" + dr["custom_assessment"];
                        score.Value = "" + dr["value"];

                        _scoreTermSubjectOriDict.Add("" + dr["ref_student_id"], new List<ESLScore>());

                        _scoreTermSubjectOriDict["" + dr["ref_student_id"]].Add(score);
                    }
                    else
                    {
                        ESLScore score = new ESLScore();

                        score.ID = "" + dr["uid"];
                        score.RefCourseID = "" + dr["ref_course_id"];
                        score.RefStudentID = "" + dr["ref_student_id"];
                        score.RefTeacherID = "" + dr["ref_teacher_id"];
                        score.RefScAttendID = "" + dr["ref_sc_attend_id"];
                        score.Term = "" + dr["term"];
                        score.Subject = "" + dr["subject"];
                        score.Assessment = "" + dr["assessment"];
                        score.Custom_Assessment = "" + dr["custom_assessment"];
                        score.Value = "" + dr["value"];

                        _scoreTermSubjectOriDict["" + dr["ref_student_id"]].Add(score);
                    }
                }
            }


            #endregion


            _worker.ReportProgress(40, "取得系統課程評量成績...");


            #region 取得學生 舊有評量 Exam 成績 

            List<string> scatIDStringList = new List<string>();

            foreach (SCAttendRecord sca in _scatList)
            {
                scatIDStringList.Add(sca.ID);
            }

            string scattendIDs = string.Join(",", scatIDStringList);

            //抓取目前所選取ESL  課程、評量，且其修課學生的 term、subject 成績
            query = @"SELECT 
                    sce_take.id
                    ,ref_sc_attend_id
                    ,ref_student_id
                    ,ref_course_id
                    ,ref_exam_id
                    ,sce_take.score
                    ,sce_take.create_date
                    ,sce_take.extension
                    FROM sce_take              
                    LEFT JOIN sc_attend ON sc_attend.id = sce_take.ref_sc_attend_id
                    WHERE ref_sc_attend_id IN( " + scattendIDs + ") " +
                    "AND ref_exam_id = " + "'" + target_exam_id + "'";


            dt = qh.Select(query);

            // 整理 既有的成績資料 
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    SCETakeESLRecord scetRecord = new SCETakeESLRecord();

                    scetRecord.ID = "" + dr["id"];
                    scetRecord.RefSCAttendID = "" + dr["ref_sc_attend_id"];
                    scetRecord.RefStudentID = "" + dr["ref_student_id"];
                    scetRecord.RefCourseID = "" + dr["ref_course_id"];
                    scetRecord.RefExamID = "" + dr["ref_exam_id"];

                    decimal d;

                    decimal.TryParse("" + dr["score"], out d);

                    scetRecord.Score = d;
                    scetRecord.Extensions = "" + dr["extension"];


                    if (!_scetESLDict.ContainsKey(scetRecord.RefStudentID))
                    {
                        _scetESLDict.Add(scetRecord.RefStudentID, new List<SCETakeESLRecord>());

                        _scetESLDict[scetRecord.RefStudentID].Add(scetRecord);
                    }
                    else
                    {
                        _scetESLDict[scetRecord.RefStudentID].Add(scetRecord);
                    }
                }
            }


            #endregion


            #region 換算ESL 成績項目 每一個的權重 、計算評量成績分數種類(定期、平時)
            foreach (string courseID in _scoreTemplateDict.Keys)
            {
                string key_subject = "";
                string key_assessment = "";

                foreach (Term t in _scoreTemplateDict[courseID])
                {
                    t.SubjectTotalWeight = 0;

                    foreach (Subject s in t.SubjectList)
                    {
                        s.AssessmentTotalWeight = 0;

                        int ratio_subject = 0;

                        foreach (Assessment a in s.AssessmentList)
                        {
                            if (a.Type == "Score") // 只取分數型成績
                            {
                                key_assessment = courseID + "_" + t.Name + "_" + s.Name + "_" + a.Name;

                                // 2019/02/01 穎驊更新， ESＬ　寒假優化，assessment　直接對應 term 成績計算
                                int ratio_assessment = int.Parse(a.Weight);

                                _scoreRatioDict.Add(key_assessment, ratio_assessment);

                                ratio_subject += ratio_assessment;

                                _scoreExamScoreTypeDict.Add(key_assessment, a.ExamScoreType); //整理分數計算類別， 以利換回系統知道成績對應(定期、平時)
                            }
                        }

                        key_subject = courseID + "_" + t.Name + "_" + s.Name;

                        _scoreRatioDict.Add(key_subject, ratio_subject);
                    }
                }
            }
            #endregion

            _worker.ReportProgress(50, "計算ESL結構成績...");

            #region ESL 成績計算
            //計算成績 依照比例 等量換算成分數 儲存  (Subject 成績)
            foreach (string studentId in _scoreAssessmentOriDict.Keys)
            {
                string key_score_assessment = "";

                string key_subject = "";
                string key_term = "";

                decimal subject_score_partial;

                foreach (ESLScore score in _scoreAssessmentOriDict[studentId])
                {
                    key_score_assessment = score.RefCourseID + "_" + score.Term + "_" + score.Subject + "_" + score.Assessment; // 查分數比例的KEY 這些就夠

                    key_subject = score.RefCourseID + "_" + score.RefStudentID + "_" + score.Term + "_" + score.Subject; // 寫給學生的Subject 成績 還必須要有 studentID 才有獨立性

                    if (_scoreRatioDict.ContainsKey(key_score_assessment))
                    {
                        decimal assementScore;
                        if (decimal.TryParse(score.Value, out assementScore))
                        {
                            //subject_score_partial = Math.Round(assementScore * _scoreRatioDict[key_score_assessment], 2, MidpointRounding.ToEven); // 四捨五入到第二位

                            // 2018/11/13 穎驊修正， 由於 康橋驗算後，發現在在小數後兩位有精度的問題，
                            // 在此統一在結算 term 為止 之前不會做任何的 四捨五入。
                            // 2019/02/26 穎驊 依據ESL 寒假優化項目， 新增了 學生個別成績 比例的計算
                            // 如果 學生該筆成績 有獨立的  Ratio 權重設定，以該Ratio 權重優先
                            if (score.Ratio != null)
                            {
                                subject_score_partial = assementScore * decimal.Parse(score.Ratio + "");
                            }
                            else
                            {
                                subject_score_partial = assementScore * _scoreRatioDict[key_score_assessment];
                            }


                            // 處理 subject分母
                            if (!_scoreRatioTotalDict.ContainsKey(key_subject))
                            {
                                if (score.Ratio != null)
                                {
                                    _scoreRatioTotalDict.Add(key_subject, decimal.Parse(score.Ratio + ""));
                                }
                                else
                                {
                                    _scoreRatioTotalDict.Add(key_subject, _scoreRatioDict[key_score_assessment]);
                                }
                            }
                            else
                            {
                                if (score.Ratio != null)
                                {
                                    _scoreRatioTotalDict[key_subject] += decimal.Parse(score.Ratio + "");
                                }
                                else
                                {
                                    _scoreRatioTotalDict[key_subject] += _scoreRatioDict[key_score_assessment];
                                }
                            }

                            // 處理 subject分母(定期、平時)
                            if (_scoreExamScoreTypeDict.ContainsKey(key_score_assessment))
                            {
                                if (!_scoreRatioTotalDict.ContainsKey(key_subject + "_" + _scoreExamScoreTypeDict[key_score_assessment]))
                                {
                                    if (score.Ratio != null)
                                    {
                                        _scoreRatioTotalDict.Add(key_subject + "_" + _scoreExamScoreTypeDict[key_score_assessment], decimal.Parse(score.Ratio + ""));
                                    }
                                    else
                                    {
                                        _scoreRatioTotalDict.Add(key_subject + "_" + _scoreExamScoreTypeDict[key_score_assessment], _scoreRatioDict[key_score_assessment]);
                                    }
                                }
                                else
                                {
                                    if (score.Ratio != null)
                                    {
                                        _scoreRatioTotalDict[key_subject + "_" + _scoreExamScoreTypeDict[key_score_assessment]] += decimal.Parse(score.Ratio + "");
                                    }
                                    else
                                    {
                                        _scoreRatioTotalDict[key_subject + "_" + _scoreExamScoreTypeDict[key_score_assessment]] += _scoreRatioDict[key_score_assessment];
                                    }
                                }
                            }



                            if (!_subjectScoreDict.ContainsKey(key_subject))
                            {
                                ESLScore subjectScore = new ESLScore();

                                subjectScore.RefCourseID = score.RefCourseID;
                                subjectScore.RefStudentID = score.RefStudentID;
                                subjectScore.RefTeacherID = score.RefTeacherID;
                                subjectScore.RefScAttendID = score.RefScAttendID;
                                subjectScore.Term = score.Term;
                                subjectScore.Subject = score.Subject;
                                subjectScore.Score = subject_score_partial;

                                _subjectScoreDict.Add(key_subject, subjectScore);
                            }
                            else
                            {
                                _subjectScoreDict[key_subject].Score += subject_score_partial;
                            }

                            // 另外存放 算成定期、平時 評量的分數
                            if (_scoreExamScoreTypeDict.ContainsKey(key_score_assessment))
                            {
                                if (!_subjectScoreDict.ContainsKey(key_subject + "_" + _scoreExamScoreTypeDict[key_score_assessment]))
                                {
                                    ESLScore subjectScore = new ESLScore();

                                    subjectScore.RefCourseID = score.RefCourseID;
                                    subjectScore.RefStudentID = score.RefStudentID;
                                    subjectScore.RefTeacherID = score.RefTeacherID;
                                    subjectScore.RefScAttendID = score.RefScAttendID;
                                    subjectScore.Term = score.Term;
                                    subjectScore.Subject = score.Subject;

                                    if (_scoreExamScoreTypeDict[key_score_assessment].Contains("定期"))
                                    {
                                        subjectScore.Score = subject_score_partial;
                                    }
                                    else
                                    {
                                        // 2018/12/17 穎驊與恩正討論後，恩正說，為了減少誤差，正確算出成績，因此平時成績用反推的                                    
                                        subjectScore.Score = 0;
                                    }


                                    _subjectScoreDict.Add(key_subject + "_" + _scoreExamScoreTypeDict[key_score_assessment], subjectScore);
                                }
                                else
                                {
                                    if (_scoreExamScoreTypeDict[key_score_assessment].Contains("定期"))
                                    {
                                        _subjectScoreDict[key_subject + "_" + _scoreExamScoreTypeDict[key_score_assessment]].Score += subject_score_partial;
                                    }
                                    else
                                    {
                                        // 2018/12/17 穎驊與恩正討論後，恩正說，為了減少誤差，正確算出成績，因此平時成績用反推的                                                                         
                                    }
                                }
                            }
                        }
                        else
                        {
                            //assementScore = 0; // 轉失敗(可能沒有輸入)，當0 分 >> 不可以!!!
                        }
                    }
                }
            }


            foreach (string key_subject in _subjectScoreDict.Keys)
            {
                ESLScore subjectScore = _subjectScoreDict[key_subject];

                string key_score_subject = subjectScore.RefCourseID + "_" + subjectScore.RefStudentID + "_" + subjectScore.Term + "_" + subjectScore.Subject; // 寫給學生的Subject 成績 還必須要有 studentID 才有獨立性

                string key_term = subjectScore.RefCourseID + "_" + subjectScore.RefStudentID + "_" + subjectScore.Term; // 寫給學生的Term 成績 還必須要有 studentID 才有獨立性

                decimal term_score_partial;

                if (_scoreRatioTotalDict.ContainsKey(key_score_subject))
                {
                    term_score_partial = subjectScore.Score;

                    // 處理 term 分母
                    if (!key_subject.Contains("定期") && !key_subject.Contains("平時"))
                    {
                        if (!_scoreRatioTotalDict.ContainsKey(key_term))
                        {
                            _scoreRatioTotalDict.Add(key_term, _scoreRatioTotalDict[key_score_subject]);
                        }
                        else
                        {
                            _scoreRatioTotalDict[key_term] += _scoreRatioTotalDict[key_score_subject];
                        }
                    }

                    // 處理 term (定期)分母 
                    if (key_subject.Contains("定期"))
                    {
                        if (!_scoreRatioTotalDict.ContainsKey(key_term + "_定期"))
                        {
                            _scoreRatioTotalDict.Add(key_term + "_定期", _scoreRatioTotalDict[key_score_subject + "_定期"]);
                        }
                        else
                        {
                            _scoreRatioTotalDict[key_term + "_定期"] += _scoreRatioTotalDict[key_score_subject + "_定期"];
                        }
                    }

                    // 處理 term (平時)分母
                    if (key_subject.Contains("平時"))
                    {
                        if (!_scoreRatioTotalDict.ContainsKey(key_term + "_平時"))
                        {
                            _scoreRatioTotalDict.Add(key_term + "_平時", _scoreRatioTotalDict[key_score_subject + "_平時"]);
                        }
                        else
                        {
                            _scoreRatioTotalDict[key_term + "_平時"] += _scoreRatioTotalDict[key_score_subject + "_平時"];
                        }

                    }

                    // 一般的term 成績
                    if (!key_subject.Contains("定期") && !key_subject.Contains("平時"))
                    {
                        if (!_termScoreDict.ContainsKey(key_term))
                        {
                            ESLScore termScore = new ESLScore();

                            termScore.RefCourseID = subjectScore.RefCourseID;
                            termScore.RefStudentID = subjectScore.RefStudentID;
                            termScore.RefTeacherID = subjectScore.RefTeacherID;
                            termScore.RefScAttendID = subjectScore.RefScAttendID;
                            termScore.Term = subjectScore.Term;
                            termScore.Score = term_score_partial;

                            _termScoreDict.Add(key_term, termScore);
                        }
                        else
                        {
                            _termScoreDict[key_term].Score += term_score_partial;
                        }
                    }

                    // 定期的term 成績
                    if (key_subject.Contains("定期"))
                    {
                        if (!_termScoreDict.ContainsKey(key_term + "_定期"))
                        {
                            ESLScore termScore = new ESLScore();

                            termScore.RefCourseID = subjectScore.RefCourseID;
                            termScore.RefStudentID = subjectScore.RefStudentID;
                            termScore.RefTeacherID = subjectScore.RefTeacherID;
                            termScore.RefScAttendID = subjectScore.RefScAttendID;
                            termScore.Term = subjectScore.Term;
                            termScore.Score = term_score_partial;

                            _termScoreDict.Add(key_term + "_定期", termScore);
                        }
                        else
                        {
                            _termScoreDict[key_term + "_定期"].Score += term_score_partial;
                        }
                    }

                    // 平時的term 成績
                    if (key_subject.Contains("平時"))
                    {
                        if (!_termScoreDict.ContainsKey(key_term + "_平時"))
                        {
                            ESLScore termScore = new ESLScore();

                            termScore.RefCourseID = subjectScore.RefCourseID;
                            termScore.RefStudentID = subjectScore.RefStudentID;
                            termScore.RefTeacherID = subjectScore.RefTeacherID;
                            termScore.RefScAttendID = subjectScore.RefScAttendID;
                            termScore.Term = subjectScore.Term;
                            termScore.Score = term_score_partial;

                            _termScoreDict.Add(key_term + "_平時", termScore);
                        }
                        else
                        {
                            _termScoreDict[key_term + "_平時"].Score += term_score_partial;
                        }
                    }
                }
            }



            // 計算Term成績後，現在將各自加權後的成績除以各自的的總權重
            foreach (KeyValuePair<string, ESLScore> score in _termScoreDict)
            {
                // 各課程  自己的 四捨五入 精度 小數位
                int _decimalPlace = _scoreDecimalPlaceDict[score.Value.RefCourseID];

                // 一般的 term 成績
                if (!score.Key.Contains("定期") && !score.Key.Contains("平時"))
                {
                    string ratioTotalTermKey = score.Value.RefCourseID + "_" + score.Value.RefStudentID + "_" + score.Value.Term;

                    _termScoreDict[score.Key].Score = Math.Round(_termScoreDict[score.Key].Score / (_scoreRatioTotalDict[ratioTotalTermKey]), _decimalPlace, MidpointRounding.AwayFromZero);
                }
                // 定期的term 成績
                if (score.Key.Contains("定期"))
                {
                    string ratioTotalTermKey = score.Value.RefCourseID + "_" + score.Value.RefStudentID + "_" + score.Value.Term + "_定期";

                    _termScoreDict[score.Key].Score = Math.Round(_termScoreDict[score.Key].Score / (_scoreRatioTotalDict[ratioTotalTermKey]), _decimalPlace, MidpointRounding.AwayFromZero);
                }
            }

            // 計算Term(一般、定期)成績後，
            // 2018/12/17 穎驊與恩正討論後，恩正說，為了減少誤差，正確算出成績，因此平時成績用反推的  
            foreach (KeyValuePair<string, ESLScore> score in _termScoreDict)
            {
                // 各課程  自己的 四捨五入 精度 小數位
                int _decimalPlace = _scoreDecimalPlaceDict[score.Value.RefCourseID];

                // term 成績
                if (!score.Key.Contains("定期") && !score.Key.Contains("平時"))
                {
                    string ratioTotalTermKey = score.Value.RefCourseID;

                    // 2018/12/17 穎驊與恩正討論後，恩正說，為了減少誤差，正確算出成績，因此平時成績用反推的  (Term 一般 -Term 定期 = Term 平時)
                    if (_termScoreDict.ContainsKey(score.Key + "_平時"))
                    {
                        if (_termScoreDict.ContainsKey(score.Key + "_定期"))
                        {
                            // 定期 加 平時 的權重 固定為 100
                            _termScoreDict[score.Key + "_平時"].Score = Math.Round(((_termScoreDict[score.Key].Score * 100 - (_termScoreDict[score.Key + "_定期"].Score) * (_scoreRatioTotalDict[ratioTotalTermKey + "_中文系統評分樣版定期比例"])) / (_scoreRatioTotalDict[ratioTotalTermKey + "_中文系統評分樣版平時比例"])), _decimalPlace, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            _termScoreDict[score.Key + "_平時"].Score = Math.Round(_termScoreDict[score.Key].Score, _decimalPlace, MidpointRounding.AwayFromZero);
                        }
                    }
                }
            }

            // 計算Subject成績後，現在將各自加權後的成績除以各自的的總權重
            foreach (KeyValuePair<string, ESLScore> score in _subjectScoreDict)
            {
                // 一般的 subject 成績
                if (!score.Key.Contains("定期") && !score.Key.Contains("平時"))
                {
                    string ratioTotalSubjectKey = score.Value.RefCourseID + "_" + score.Value.RefStudentID + "_" + score.Value.Term + "_" + score.Value.Subject;

                    // 2018/11/13 穎驊修正， 由於 康橋驗算後，發現在在小數後兩位有精度的問題，
                    // 在此統一在結算 term 為止 之前不會做任何的 四捨五入。
                    _subjectScoreDict[score.Key].Score = _subjectScoreDict[score.Key].Score / _scoreRatioTotalDict[ratioTotalSubjectKey];
                }
            }

            // 2018/11/13 穎驊更新， 已經計算完 Term 成績，現在可以把 Subject 成績 四捨五入
            foreach (KeyValuePair<string, ESLScore> score in _subjectScoreDict)
            {
                // 各課程  自己的 四捨五入 精度 小數位
                int _decimalPlace = _scoreDecimalPlaceDict[score.Value.RefCourseID];

                _subjectScoreDict[score.Key].Score = Math.Round(_subjectScoreDict[score.Key].Score, _decimalPlace, MidpointRounding.AwayFromZero);
            }

            // 2018/12/28 穎驊更新，成績都計算完後， trim 掉 多餘的 0， 像是 85.0 >> 85 、 70.0 >> 70 、76.1234000 >> 76.1234
            foreach (KeyValuePair<string, ESLScore> score in _subjectScoreDict)
            {
                _subjectScoreDict[score.Key].Score = decimal.Parse(_subjectScoreDict[score.Key].Score.ToString("0.#########"));
            }

            // 以 studentID 為 key 整理 學生subject成績 至_scorefinalDict
            foreach (KeyValuePair<string, ESLScore> score in _subjectScoreDict)
            {
                if (!score.Key.Contains("定期") && !score.Key.Contains("平時"))
                {
                    if (!_scorefinalDict.ContainsKey(score.Value.RefStudentID))
                    {
                        _scorefinalDict.Add(score.Value.RefStudentID, new List<ESLScore>());

                        _scorefinalDict[score.Value.RefStudentID].Add(score.Value);
                    }
                    else
                    {
                        _scorefinalDict[score.Value.RefStudentID].Add(score.Value);
                    }
                }
            }

            // 以 studentID 為 key 整理 學生term 成績 至 _scorefinalDict
            foreach (KeyValuePair<string, ESLScore> score in _termScoreDict)
            {
                if (!score.Key.Contains("定期") && !score.Key.Contains("平時"))
                {
                    if (!_scorefinalDict.ContainsKey(score.Value.RefStudentID))
                    {
                        _scorefinalDict.Add(score.Value.RefStudentID, new List<ESLScore>());

                        _scorefinalDict[score.Value.RefStudentID].Add(score.Value);
                    }
                    else
                    {
                        _scorefinalDict[score.Value.RefStudentID].Add(score.Value);
                    }
                }
            }
            #endregion

            List<ESLScore> updateESLscoreList = new List<ESLScore>(); // 最後要update ESLscoreList
            List<ESLScore> insertESLscoreList = new List<ESLScore>(); // 最後要indert ESLscoreList

            foreach (string studentID in _scorefinalDict.Keys)
            {
                foreach (ESLScore scoreFinal in _scorefinalDict[studentID])
                {
                    string scoreKey = scoreFinal.RefCourseID + "_" + scoreFinal.Subject + "_" + scoreFinal.Term;

                    if (_scoreTermSubjectOriDict.ContainsKey(studentID))
                    {
                        foreach (ESLScore scoreOri in _scoreTermSubjectOriDict[studentID])
                        {
                            // update 分數
                            if (scoreKey == scoreOri.RefCourseID + "_" + scoreOri.Subject + "_" + scoreOri.Term)
                            {
                                scoreOri.Score = scoreFinal.Score;
                                updateESLscoreList.Add(scoreOri);
                            }
                        }

                        // 假若就分數沒有任何一項對的起來 就是 insert 分數
                        if (!_scoreTermSubjectOriDict[studentID].Any(s => s.RefCourseID + "_" + s.Subject + "_" + s.Term == scoreKey))
                        {
                            insertESLscoreList.Add(scoreFinal);
                        }


                    }
                    //  insert 分數
                    else
                    {
                        insertESLscoreList.Add(scoreFinal);
                    }
                }
            }



            ////2018/6/14 先暫時這樣整理，之後會想要用 scoreUpsertDict ，資料會比較齊
            //foreach (ESLScore score in _subjectScoreDict.Values)
            //{
            //    _eslscoreList.Add(score);
            //}

            //foreach (ESLScore score in _termScoreDict.Values)
            //{
            //    _eslscoreList.Add(score);
            //}


            _worker.ReportProgress(80, "轉換ESL成績 為評量成績");

            #region 換算Exam 成績
            //List<SCETakeRecord> updateList = new List<SCETakeRecord>();
            //List<SCETakeRecord> insertList = new List<SCETakeRecord>();

            #region 舊的 使用 K12 API 方法
            // 更新舊評量分數
            //foreach (string studentID in _scetDict.Keys)
            //{
            //    foreach (K12.Data.SCETakeRecord sce in _scetDict[studentID])
            //    {
            //        foreach (ESLScore score in _termScoreDict.Values)
            //        {
            //            if (sce.RefCourseID == score.RefCourseID && sce.RefStudentID == score.RefStudentID && GetScore(sce) != "" + score.Score)
            //            {
            //                SetScore(sce, "" + score.Score);

            //                updateList.Add(sce);
            //            }
            //        }
            //    }
            //}

            //// 新增 評量分數
            //foreach (SCAttendRecord sca in _scatList)
            //{
            //    if (_scetDict.ContainsKey(sca.RefStudentID))
            //    {
            //        // 有學生修課紀錄，卻在該試別 沒有舊評量成績 就是本次要新增的項目
            //        if (!_scetDict[sca.RefStudentID].Any(s => s.RefSCAttendID == sca.ID))
            //        {
            //            foreach (ESLScore score in _termScoreDict.Values)
            //            {
            //                if (sca.RefCourseID == score.RefCourseID && sca.RefStudentID == score.RefStudentID)
            //                {
            //                    SCETakeRecord sce = new SCETakeRecord();
            //                    sce.RefSCAttendID = sca.ID;
            //                    sce.RefExamID = target_exam_id;
            //                    sce.RefStudentID = sca.RefStudentID;
            //                    sce.RefCourseID = sca.RefCourseID;
            //                    SetScore(sce, "" + score.Score);
            //                    insertList.Add(sce);
            //                }
            //            }
            //        }

            //    }
            //    else
            //    {
            //        foreach (ESLScore score in _termScoreDict.Values)
            //        {
            //            if (sca.RefCourseID == score.RefCourseID && sca.RefStudentID == score.RefStudentID)
            //            {
            //                SCETakeRecord sce = new SCETakeRecord();
            //                sce.RefSCAttendID = sca.ID;
            //                sce.RefExamID = target_exam_id;
            //                sce.RefStudentID = sca.RefStudentID;
            //                sce.RefCourseID = sca.RefCourseID;
            //                SetScore(sce, "" + score.Score);
            //                insertList.Add(sce);
            //            }
            //        }
            //    }
            //}


            //K12.Data.SCETake.Update(updateList);
            //K12.Data.SCETake.Insert(insertList);
            #endregion


            List<SCETakeESLRecord> updateList = new List<SCETakeESLRecord>();
            List<SCETakeESLRecord> insertList = new List<SCETakeESLRecord>();

            // 更新舊評量分數 
            foreach (string studentID in _scetESLDict.Keys)
            {
                foreach (SCETakeESLRecord sce in _scetESLDict[studentID])
                {
                    foreach (string key in _termScoreDict.Keys)
                    {
                        ESLScore score = _termScoreDict[key];

                        // 總評量成績
                        if (!key.Contains("定期") && !key.Contains("平時"))
                        {
                            if (sce.RefCourseID == score.RefCourseID && sce.RefStudentID == score.RefStudentID && "" + sce.Score != "" + score.Score)
                            {
                                sce.Score = score.Score;

                                sce.NeedUpdate = true;
                            }
                        }

                        // 定期的 term 成績
                        if (key.Contains("定期"))
                        {
                            if (sce.RefCourseID == score.RefCourseID && sce.RefStudentID == score.RefStudentID && GetScore(sce) != "" + score.Score)
                            {
                                SetScore(sce, "" + score.Score);
                                sce.NeedUpdate = true;
                            }
                        }

                        // 平時的 term 成績
                        if (key.Contains("平時"))
                        {
                            if (sce.RefCourseID == score.RefCourseID && sce.RefStudentID == score.RefStudentID && GetAssignmentScore(sce) != "" + score.Score)
                            {
                                SetAssignmentScore(sce, "" + score.Score);

                                sce.NeedUpdate = true;
                            }
                        }
                    }
                }
            }

            // 更新舊評量分數 
            foreach (string studentID in _scetESLDict.Keys)
            {
                foreach (SCETakeESLRecord sce in _scetESLDict[studentID])
                {
                    if (sce.NeedUpdate)
                    {
                        updateList.Add(sce);
                    }
                }
            }

            // 新增 評量分數
            foreach (SCAttendRecord sca in _scatList)
            {
                if (_scetESLDict.ContainsKey(sca.RefStudentID))
                {
                    // 有學生修課紀錄，卻在該試別 沒有舊評量成績 就是本次要新增的項目
                    if (!_scetESLDict[sca.RefStudentID].Any(s => s.RefSCAttendID == sca.ID))
                    {
                        SCETakeESLRecord sce = new SCETakeESLRecord();

                        foreach (string key in _termScoreDict.Keys)
                        {
                            ESLScore score = _termScoreDict[key];

                            if (!key.Contains("定期") && !key.Contains("平時") && sca.RefCourseID == score.RefCourseID && sca.RefStudentID == score.RefStudentID)
                            {
                                sce.RefSCAttendID = sca.ID;
                                sce.RefExamID = target_exam_id;
                                sce.RefStudentID = sca.RefStudentID;
                                sce.RefCourseID = sca.RefCourseID;
                                SetScore(sce, "" + score.Score);
                                sce.Score = score.Score;
                            }

                            if (key.Contains("定期") && sca.RefCourseID == score.RefCourseID && sca.RefStudentID == score.RefStudentID)
                            {
                                sce.RefSCAttendID = sca.ID;
                                sce.RefExamID = target_exam_id;
                                sce.RefStudentID = sca.RefStudentID;
                                sce.RefCourseID = sca.RefCourseID;
                                SetScore(sce, "" + score.Score);
                            }
                            if (key.Contains("平時") && sca.RefCourseID == score.RefCourseID && sca.RefStudentID == score.RefStudentID)
                            {
                                sce.RefSCAttendID = sca.ID;
                                sce.RefExamID = target_exam_id;
                                sce.RefStudentID = sca.RefStudentID;
                                sce.RefCourseID = sca.RefCourseID;
                                SetAssignmentScore(sce, "" + score.Score);
                            }
                        }
                        if (sce.RefSCAttendID != null)
                        {
                            insertList.Add(sce);
                        }

                    }

                }
                else
                {
                    SCETakeESLRecord sce = new SCETakeESLRecord();

                    foreach (string key in _termScoreDict.Keys)
                    {
                        ESLScore score = _termScoreDict[key];

                        if (!key.Contains("定期") && !key.Contains("平時") && sca.RefCourseID == score.RefCourseID && sca.RefStudentID == score.RefStudentID)
                        {
                            sce.RefSCAttendID = sca.ID;
                            sce.RefExamID = target_exam_id;
                            sce.RefStudentID = sca.RefStudentID;
                            sce.RefCourseID = sca.RefCourseID;
                            SetScore(sce, "" + score.Score);
                            sce.Score = score.Score;
                        }

                        if (key.Contains("定期") && sca.RefCourseID == score.RefCourseID && sca.RefStudentID == score.RefStudentID)
                        {
                            sce.RefSCAttendID = sca.ID;
                            sce.RefExamID = target_exam_id;
                            sce.RefStudentID = sca.RefStudentID;
                            sce.RefCourseID = sca.RefCourseID;
                            SetScore(sce, "" + score.Score);
                        }
                        if (key.Contains("平時") && sca.RefCourseID == score.RefCourseID && sca.RefStudentID == score.RefStudentID)
                        {
                            sce.RefSCAttendID = sca.ID;
                            sce.RefExamID = target_exam_id;
                            sce.RefStudentID = sca.RefStudentID;
                            sce.RefCourseID = sca.RefCourseID;
                            SetAssignmentScore(sce, "" + score.Score);
                        }
                    }
                    if (sce.RefSCAttendID != null)
                    {
                        insertList.Add(sce);
                    }
                }
            }

            //拚SQL
            // 兜資料
            List<string> examDataList = new List<string>();

            foreach (SCETakeESLRecord score in updateList)
            {
                string data = string.Format(@"
                SELECT
                    '{0}'::BIGINT AS ref_sc_attend_id
                    ,'{1}'::BIGINT AS ref_exam_id
                    ,'{2}'::DECIMAL AS score
                    ,'{3}'::TEXT AS extension                                        
                    ,'{4}'::INTEGER AS id
                    ,'UPDATE'::TEXT AS action
                ", score.RefSCAttendID, score.RefExamID, score.Score, score.Extensions, score.ID);

                examDataList.Add(data);
            }

            foreach (SCETakeESLRecord score in insertList)
            {
                string data = string.Format(@"
                SELECT
                '{0}'::BIGINT AS ref_sc_attend_id
                    ,'{1}'::BIGINT AS ref_exam_id
                    ,'{2}'::DECIMAL AS score
                    ,'{3}'::TEXT AS extension                                        
                    ,'{4}'::INTEGER AS id
                    ,'INSERT'::TEXT AS action
                ", score.RefSCAttendID, score.RefExamID, score.Score, score.Extensions, 0);

                examDataList.Add(data);
            }


            string examData = string.Join(" UNION ALL", examDataList);


            string examsql = string.Format(@"
WITH score_data_row AS(			 
                {0}     
),update_score AS(	    
    Update sce_take
    SET
        ref_sc_attend_id = score_data_row.ref_sc_attend_id
        ,ref_exam_id = score_data_row.ref_exam_id
        ,score = score_data_row.score
        ,extension = score_data_row.extension
    FROM 
        score_data_row    
    WHERE sce_take.id = score_data_row.id  
        AND score_data_row.action ='UPDATE'
    RETURNING  sce_take.* 
)
INSERT INTO sce_take(
	ref_sc_attend_id	
	,ref_exam_id
    ,score
	,extension
)	
SELECT 
	score_data_row.ref_sc_attend_id::BIGINT AS ref_sc_attend_id	
	,score_data_row.ref_exam_id::BIGINT AS ref_exam_id
    ,score_data_row.score::DECIMAL AS score
	,score_data_row.extension::TEXT AS extension		
FROM
	score_data_row
WHERE action ='INSERT'", examData);
            UpdateHelper uh = new UpdateHelper();

            if (!string.IsNullOrWhiteSpace(examData))
            {


                _worker.ReportProgress(80, "上傳成績...");

                //執行sql
                uh.Execute(examsql);
            }

            #endregion


            //拚SQL
            // 兜資料
            List<string> dataList = new List<string>();


            // 沒有新增任何成績資料，代表所選ESL 課程都沒有成績，不需執行SQL
            if (updateESLscoreList.Count + insertESLscoreList.Count != 0)
            {
                //  return;



                foreach (ESLScore score in updateESLscoreList)
                {
                    string data = string.Format(@"
                SELECT
                    '{0}'::BIGINT AS ref_student_id
                    ,'{1}'::BIGINT AS ref_course_id
                    ,'{2}'::BIGINT AS ref_teacher_id
                    ,'{3}'::BIGINT AS ref_sc_attend_id
                    ,'{4}'::TEXT AS term
                    ,{5} AS subject
                    ,'{6}'::TEXT AS value
                    ,'{7}'::INTEGER AS uid
                    ,'UPDATE'::TEXT AS action
                ", score.RefStudentID, score.RefCourseID, score.RefTeacherID, score.RefScAttendID, score.Term, score.Subject != null ? "'" + score.Subject + "' ::TEXT" : "NULL", score.Score, score.ID);

                    dataList.Add(data);
                }

                foreach (ESLScore score in insertESLscoreList)
                {
                    string data = string.Format(@"
                SELECT
                    '{0}'::BIGINT AS ref_student_id
                    ,'{1}'::BIGINT AS ref_course_id
                    ,'{2}'::BIGINT AS ref_teacher_id
                    ,'{3}'::BIGINT AS ref_sc_attend_id
                    ,'{4}'::TEXT AS term
                    ,{5} AS subject
                    ,'{6}'::TEXT AS value
                    ,{7}::INTEGER AS uid
                    ,'INSERT'::TEXT AS action
                ", score.RefStudentID, score.RefCourseID, score.RefTeacherID, score.RefScAttendID, score.Term, score.Subject != null ? "'" + score.Subject + "' ::TEXT" : "NULL", score.Score, 0);  // insert 給 uid = 0

                    dataList.Add(data);
                }

                string Data = string.Join(" UNION ALL", dataList);


                string sql = string.Format(@"
WITH score_data_row AS(			 
                {0}     
),update_score AS(	    
    Update $esl.gradebook_assessment_score
    SET
        ref_student_id = score_data_row.ref_student_id
        ,ref_course_id = score_data_row.ref_course_id
        ,ref_teacher_id = score_data_row.ref_teacher_id
        ,ref_sc_attend_id = score_data_row.ref_sc_attend_id
        ,term = score_data_row.term
        ,subject = score_data_row.subject
        ,value = score_data_row.value
        ,last_update = NOW()
    FROM 
        score_data_row    
    WHERE $esl.gradebook_assessment_score.uid = score_data_row.uid  
        AND score_data_row.action ='UPDATE'
    RETURNING  $esl.gradebook_assessment_score.* 
)
INSERT INTO $esl.gradebook_assessment_score(
	ref_student_id	
	,ref_course_id
    ,ref_teacher_id
    ,ref_sc_attend_id
	,term
	,subject
	,value
)
SELECT 
	score_data_row.ref_student_id::BIGINT AS ref_student_id	
	,score_data_row.ref_course_id::BIGINT AS ref_course_id
    ,score_data_row.ref_teacher_id::BIGINT AS ref_teacher_id
    ,score_data_row.ref_sc_attend_id::BIGINT AS ref_sc_attend_id
	,score_data_row.term::TEXT AS term	
	,score_data_row.subject::TEXT AS subject	
	,score_data_row.value::TEXT AS value	
FROM
	score_data_row
WHERE action ='INSERT'", Data);



                uh = new UpdateHelper();


                if (!string.IsNullOrWhiteSpace(sql))
                {
                    _worker.ReportProgress(90, "上傳成績...");
                    //執行sql
                    uh.Execute(sql);
                }
            }


            #region 將 Commet 組合 回寫 sce_take 相對 文字描述
            // 取得課程內本次評量有Commen
            Dictionary<string, List<string>> tmpCourseDict = new Dictionary<string, List<string>>();
            foreach (string cid in _scoreTemplateDict.Keys)
            {
                foreach (Term t in _scoreTemplateDict[cid])
                {
                    // 這次試別
                    if (t.Ref_exam_id == target_exam_id)
                    {
                        foreach (Subject s in t.SubjectList)
                        {
                            foreach (Assessment a in s.AssessmentList)
                            {
                                if (a.Type == "Comment")
                                {
                                    // 紀錄 Key
                                    string key = t.Name + "_" + s.Name + "_" + a.Name;
                                    if (!tmpCourseDict.ContainsKey(cid))
                                        tmpCourseDict.Add(cid, new List<string>());

                                    tmpCourseDict[cid].Add(key);
                                }
                            }
                        }
                    }
                }
            }
            // 取得資料修課學生 Comment 值

            query = @"SELECT 
                      $esl.gradebook_assessment_score.uid
                     ,sc_attend.ref_course_id
                     ,sc_attend.ref_student_id
                     ,$esl.gradebook_assessment_score.ref_teacher_id
                     ,$esl.gradebook_assessment_score.ref_sc_attend_id
                     ,$esl.gradebook_assessment_score.term
                     ,$esl.gradebook_assessment_score.subject
                     ,$esl.gradebook_assessment_score.assessment
                     ,$esl.gradebook_assessment_score.value
                      FROM $esl.gradebook_assessment_score 
                      LEFT JOIN sc_attend ON sc_attend.id = $esl.gradebook_assessment_score.ref_sc_attend_id" +
                  " WHERE ref_sc_attend_id IN( " + sc_attend_IDs + ") " +
                  " AND term IN(" + termNames + ")" +
                  "";

            // 比對後每位修課學生 Comment
            //ref_sc_attend_id,Comment
            Dictionary<string, List<string>> StudCommentDict = new Dictionary<string, List<string>>();
            QueryHelper qha = new QueryHelper();
            DataTable dt1 = qha.Select(query);

            foreach (DataRow dr in dt1.Rows)
            {
                string id = dr["ref_sc_attend_id"].ToString();

                string cid = dr["ref_course_id"].ToString();

                string ck = dr["term"] + "_" + dr["subject"] + "_" + dr["assessment"];

                if (tmpCourseDict.ContainsKey(cid))
                {
                    if (tmpCourseDict[cid].Contains(ck))
                    {
                        // 是 Comment
                        if (!StudCommentDict.ContainsKey(id))
                            StudCommentDict.Add(id, new List<string>());

                        if (dr["value"] != null && dr["value"].ToString() != "")
                        {
                            StudCommentDict[id].Add(dr["value"].ToString());
                        }
                    }
                }
            }

            // 回寫試
            query = @"SELECT 
                    sce_take.id
                    ,ref_sc_attend_id
                    ,ref_student_id
                    ,ref_course_id
                    ,ref_exam_id
                    ,sce_take.score
                    ,sce_take.create_date
                    ,sce_take.extension
                    FROM sce_take              
                    LEFT JOIN sc_attend ON sc_attend.id = sce_take.ref_sc_attend_id
                    WHERE ref_sc_attend_id IN( " + scattendIDs + ") " +
                   "AND ref_exam_id = " + "'" + target_exam_id + "'";

            DataTable dt2 = qha.Select(query);
            List<string> updateData = new List<string>();
            foreach (DataRow dr in dt2.Rows)
            {
                string sceid = dr["id"].ToString();
                string scid = dr["ref_sc_attend_id"].ToString();
                string extension = dr["extension"].ToString();

                if (StudCommentDict.ContainsKey(scid))
                {
                    string text = string.Join(",", StudCommentDict[scid].ToArray());

                    try
                    {
                        XElement elmRoot = null;

                        if (string.IsNullOrWhiteSpace(extension))
                        {
                            elmRoot = new XElement("Extension");
                        }
                        else
                        {
                            elmRoot = XElement.Parse(extension);
                        }
                        elmRoot.SetElementValue("Text", text);
                        string new_extension = elmRoot.ToString();
                        string updateQry = "UPDATE sce_take SET extension='" + new_extension + "' WHERE id=" + sceid;
                        updateData.Add(updateQry);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                }
            }

            try
            {
                if (updateData.Count > 0)
                {
                    UpdateHelper uh1 = new UpdateHelper();
                    uh1.Execute(updateData);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            #endregion

            _worker.ReportProgress(100, "ESL 評量成績計算完成。");

        }


        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MsgBox.Show("計算失敗!!，錯誤訊息:" + e.Error.Message);

            }
            else if (e.Cancelled)
            {
                MsgBox.Show("計算中止!!，中止訊息: 所選擇ESL課程，並無任何ESL 成績資料，請檢查。");
            }
            else
            {
                MsgBox.Show("計算完成!");
            }


        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("" + e.UserState, e.ProgressPercentage);
        }


        private string GetScore(SCETakeESLRecord sce)
        {
            string xmlStr = "<Extension>" + sce.Extensions + "</Extension>";
            XElement elmRoot = XElement.Parse(xmlStr);

            XmlElement xmlElement = null;
            XmlReader xmlReader = null;
            try
            {
                xmlReader = elmRoot.CreateReader();
                var doc = new XmlDocument();
                xmlElement = doc.ReadNode(elmRoot.CreateReader()) as XmlElement;
            }
            catch
            {
            }
            finally
            {
                if (xmlReader != null) xmlReader.Close();
            }


            XmlElement elem = xmlElement.SelectSingleNode("Extension/Extension/Score") as XmlElement;

            string score = elem == null ? string.Empty : elem.InnerText;

            return score;
        }

        private string GetAssignmentScore(SCETakeESLRecord sce)
        {
            string xmlStr = "<Extension>" + sce.Extensions + "</Extension>";
            XElement elmRoot = XElement.Parse(xmlStr);

            XmlElement xmlElement = null;
            XmlReader xmlReader = null;
            try
            {
                xmlReader = elmRoot.CreateReader();
                var doc = new XmlDocument();
                xmlElement = doc.ReadNode(elmRoot.CreateReader()) as XmlElement;
            }
            catch
            {
            }
            finally
            {
                if (xmlReader != null) xmlReader.Close();
            }


            XmlElement elem = xmlElement.SelectSingleNode("Extension/Extension/AssignmentScore") as XmlElement;

            string score = elem == null ? string.Empty : elem.InnerText;

            return score;
        }


        private void SetScore(SCETakeESLRecord sce, string score)
        {
            if (sce.Extensions != null)
            {
                string xmlStr = "<Extension>" + sce.Extensions + "</Extension>";
                XElement elmRoot = XElement.Parse(xmlStr);

                XmlElement xmlElement = null;
                XmlReader xmlReader = null;
                try
                {
                    xmlReader = elmRoot.CreateReader();
                    var doc = new XmlDocument();
                    xmlElement = doc.ReadNode(elmRoot.CreateReader()) as XmlElement;
                }
                catch
                {
                }
                finally
                {
                    if (xmlReader != null) xmlReader.Close();
                }

                XmlElement elem = xmlElement.SelectSingleNode("Extension/Score") as XmlElement;

                decimal d;
                decimal.TryParse(score, out d);

                if (elem != null)
                {
                    elem.InnerText = d + "";
                }
                else
                {
                    XmlElement elemScore = xmlElement.OwnerDocument.CreateElement("Score");

                    elemScore.InnerText = d + "";

                    xmlElement.SelectSingleNode("Extension").AppendChild(elemScore);
                }

                sce.Extensions = (xmlElement.InnerXml);
            }
            else
            {
                sce.Extensions = "<Extension><Score>" + score + "</Score><AssignmentScore/><Text/><Effort/></Extension> ";
            }

        }

        private void SetAssignmentScore(SCETakeESLRecord sce, string score)
        {
            if (sce.Extensions != null)
            {
                string xmlStr = "<Extension>" + sce.Extensions + "</Extension>";
                XElement elmRoot = XElement.Parse(xmlStr);

                XmlElement xmlElement = null;
                XmlReader xmlReader = null;
                try
                {
                    xmlReader = elmRoot.CreateReader();
                    var doc = new XmlDocument();
                    xmlElement = doc.ReadNode(elmRoot.CreateReader()) as XmlElement;
                }
                catch
                {
                }
                finally
                {
                    if (xmlReader != null) xmlReader.Close();
                }

                XmlElement elem = xmlElement.SelectSingleNode("Extension/AssignmentScore") as XmlElement;

                decimal d;
                decimal.TryParse(score, out d);

                if (elem != null)
                {
                    elem.InnerText = d + "";
                }
                else
                {
                    XmlElement elemAssignmentScore = xmlElement.OwnerDocument.CreateElement("AssignmentScore");

                    elemAssignmentScore.InnerText = d + "";

                    xmlElement.SelectSingleNode("Extension").AppendChild(elemAssignmentScore);
                }

                sce.Extensions = (xmlElement.InnerXml);
            }
            else
            {
                sce.Extensions = "<Extension><Score/><AssignmentScore>" + score + "</AssignmentScore><Text/><Effort/></Extension> ";
            }

        }







    }



}
