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
using Aspose.Words;
using System.Xml.Linq;

namespace ESL_System.Form
{
    public partial class PrintESLReportForm : BaseForm
    {

        private List<string> _courseIDList;

        private List<K12.Data.CourseRecord> _eslCouseList;

        private List<string> _refAssessmentSetupIDList;

        private BackgroundWorker _worker;
        
        // 依照樣板編號 分類 各自的 樣板
        private Dictionary<string, Document> _documentDict;

        // 儲放學生ESL 成績的dict 其結構為 <studentID_courseID,<scoreKey,scoreValue>
        private Dictionary<string, Dictionary<string, string>> _scoreDict = new Dictionary<string, Dictionary<string, string>>();

        // 儲放ESL 成績單 科目、比重設定 的dict 其結構為 <courseID,<key,value>>
        private Dictionary<string, Dictionary<string, string>> _itemDict = new Dictionary<string, Dictionary<string, string>>();

        // 儲放每一個評分樣版 對應的 功能變數欄位 <assessmentSetupID,DataTable>
        private Dictionary<string, DataTable> _assessmentSetupDataTableDict = new Dictionary<string, DataTable>();
        
        // 評分樣版 id 對應的修課課程id 對照
        private Dictionary<string, string> _assessmentSetupIDPairDict = new Dictionary<string, string>();

        // 修課課程id 對應的評分樣版 id 對照
        private Dictionary<string, string> _courseIDPairDict = new Dictionary<string, string>();

        // 紀錄成績 為 指標型indicator 的 key值 ， 作為對照 key 為 courseID_termName_subjectName_assessment_Name
        private List<string> _indicatorList = new List<string>();

        // 紀錄成績 為 評語型comment 的 key值
        private List<string> _commentList = new List<string>();

        // 期中/期末/ 學期
        private string _examType;


        public PrintESLReportForm(List<string> courseIDList)
        {
            InitializeComponent();
            _courseIDList = courseIDList;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            buttonX1.Enabled = false; // 關閉按鈕
            // 驗證完畢，開始列印報表
            PrintReport(_courseIDList);
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CheckCalculateTermForm_Load(object sender, EventArgs e)
        {
            comboBoxEx1.Items.Add("期中");
            comboBoxEx1.Items.Add("期末");
            comboBoxEx1.Items.Add("學期");

            comboBoxEx1.SelectedIndex = 0;
        }


        // 列印 ESL 報表
        private void PrintReport(List<string> courseIDList)
        {
            _examType = comboBoxEx1.Text;

            _worker = new BackgroundWorker();
            _worker.DoWork += new DoWorkEventHandler(Worker_DoWork);
            _worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Worker_RunWorkerCompleted);
            _worker.ProgressChanged += new ProgressChangedEventHandler(Worker_ProgressChanged);
            _worker.WorkerReportsProgress = true;
           
            _worker.RunWorkerAsync();

        }


        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            // 處理等第
            DegreeMapper dm = new DegreeMapper();

            _worker.ReportProgress(0, "開始列印 ESL報表...");

            #region 取得課程成績單 設定樣板
            _eslCouseList = new List<K12.Data.CourseRecord>();

            _eslCouseList = K12.Data.Course.SelectByIDs(_courseIDList);

            _refAssessmentSetupIDList = new List<string>();

            foreach (K12.Data.CourseRecord courseRecord in _eslCouseList)
            {
                if (!_refAssessmentSetupIDList.Contains("'" + courseRecord.RefAssessmentSetupID + "'"))
                {
                    _refAssessmentSetupIDList.Add("'" + courseRecord.RefAssessmentSetupID + "'");
                }

                if (!_courseIDPairDict.ContainsKey(courseRecord.ID))
                {
                    _courseIDPairDict.Add(courseRecord.ID, courseRecord.RefAssessmentSetupID);
                }

                if (!_assessmentSetupIDPairDict.ContainsKey(courseRecord.RefAssessmentSetupID))
                {
                    _assessmentSetupIDPairDict.Add(courseRecord.RefAssessmentSetupID, courseRecord.ID);
                }
            }

            string assessmentSetupIDs = string.Join(",", _refAssessmentSetupIDList);


            FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

            _worker.ReportProgress(0, "取得課程成績單設定樣板...");


            string qry = "ref_exam_template_id IN (" + assessmentSetupIDs + ") and schoolyear='" + K12.Data.School.DefaultSchoolYear + "' and semester ='" + K12.Data.School.DefaultSemester + "' and exam ='" + _examType + "'";

            List<UDT_ReportTemplate> configures = _AccessHelper.Select<UDT_ReportTemplate>(qry);

            _documentDict = new Dictionary<string, Document>();

            foreach (UDT_ReportTemplate templateconfig in configures)
            {
                if (!_documentDict.ContainsKey(templateconfig.Ref_exam_Template_ID))
                {
                    Document _doc = new Document();

                    templateconfig.Decode(); // 將 stream 轉成 Word

                    _doc = templateconfig.Template;

                    _documentDict.Add(templateconfig.Ref_exam_Template_ID, _doc);

                }
            }
            #endregion


            #region 取得修課學生、 並做整理
            List<K12.Data.SCAttendRecord> scList = K12.Data.SCAttend.SelectByCourseIDs(_courseIDList);

            List<string> studentIDList = new List<string>();


            foreach (K12.Data.SCAttendRecord scr in scList)
            {
                studentIDList.Add(scr.Student.ID);

                // 建立成績整理 Dict ，[studentID_courseID,[scoreKey,scoreID]]
                _scoreDict.Add(scr.Student.ID + "_" + scr.Course.ID, new Dictionary<string, string>());
            } 
            #endregion

            // 建立功能變數對照
            CreateFieldTemplate();

            #region 取得、整理ESL成績
            _worker.ReportProgress(20, "取得ESL課程成績");


            int progress = 80;
            decimal per = (decimal)(100 - progress) / scList.Count;
            int count = 0;

            string course_ids = string.Join("','", _courseIDList);

            string student_ids = string.Join("','", studentIDList);

            string sql = "SELECT * FROM $esl.gradebook_assessment_score WHERE ref_course_id IN ('" + course_ids + "') AND ref_student_id IN ('" + student_ids + "') ORDER BY last_update "; // 2018/6/21 通通都抓了，因為一張成績單上資訊，不只Final的

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(sql);


            foreach (DataRow row in dt.Rows)
            {
                string termWord = "" + row["term"];
                string subjectWord = "" + row["subject"];
                string assessmentWord = "" + row["assessment"];

                string id = "" + row["ref_student_id"] + "_" + row["ref_course_id"];

                // 有教師自訂的子項目成績就跳掉 不處理
                if ("" + row["custom_assessment"] != "")
                {
                    continue;
                }

                // 要設計一個模式 處理 三種成績

                // 項目都有，為assessment 成績
                if (termWord != "" && "" + subjectWord != "" && "" + assessmentWord != "")
                {
                    if (_scoreDict.ContainsKey(id))
                    {
                        // 指標型成績
                        if (_indicatorList.Contains("" + row["ref_course_id"] + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_')))
                        {
                            string key = "評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "指標";
                            if (_scoreDict[id].ContainsKey(key))
                            {
                                _scoreDict[id][key] = "" + row["value"]; //重覆項目，後來時間的蓋過前面
                            }
                            else
                            {
                                _scoreDict[id].Add(key, "" + row["value"]);
                            }

                        }
                        // 評語型成績
                        else if (_commentList.Contains("" + row["ref_course_id"] + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_')))
                        {
                            string key = "評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "評語";
                            if (_scoreDict[id].ContainsKey(key))
                            {
                                _scoreDict[id][key] = "" + row["value"]; //重覆項目，後來時間的蓋過前面
                            }
                            else
                            {
                                _scoreDict[id].Add(key, "" + row["value"]);
                            }
                        }
                        // 分數型成績
                        else
                        {
                            string key = "評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數";
                            if (_scoreDict[id].ContainsKey(key))
                            {
                                _scoreDict[id][key] = "" + row["value"]; //重覆項目，後來時間的蓋過前面
                            }
                            else
                            {
                                _scoreDict[id].Add(key, "" + row["value"]);
                            }

                        }
                    }
                }

                // 沒有assessment，為subject 成績
                if (termWord != "" && "" + subjectWord != "" && "" + assessmentWord == "")
                {
                    if (_scoreDict.ContainsKey(id))
                    {
                        _scoreDict[id].Add("評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數", "" + row["value"]);

                    }
                }
                // 沒有assessment、subject，為term 成績
                if (termWord != "" && "" + subjectWord == "" && "" + assessmentWord == "")
                {

                    if (_scoreDict.ContainsKey(id))
                    {
                        _scoreDict[id].Add("評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數", "" + row["value"]);

                    }
                }

            }



            // 課程學期成績
            string sqlSemesterCourseScore = @"SELECT
sc_attend.ref_student_id
,sc_attend.ref_course_id
,exam_template.name
,course.domain
,course.subject
,sc_attend.score
FROM sc_attend
LEFT JOIN course ON course.id = sc_attend.ref_course_id
LEFT JOIN exam_template ON exam_template.id =course.ref_exam_template_id
WHERE course.id IN ('" + course_ids + "') " +
"AND sc_attend.ref_student_id IN ('" + student_ids + "')" +
"ORDER BY ref_student_id,domain,subject";

            DataTable dtSemesterCourseScore = qh.Select(sqlSemesterCourseScore);

            foreach (DataRow row in dtSemesterCourseScore.Rows)
            {
                string id = "" + row["ref_student_id"] + "_" + row["ref_course_id"];

                string templateWord = "" + row["name"];
                string domainWord = "" + row["domain"];
                string subjectWord = "" + row["subject"];

                string score = "" + row["score"]; // 課程學期成績

               
                if (_scoreDict.ContainsKey(id))
                {
                    #region 跟樣板的功能變數
                    // 理論上一學期上 一個學生 只會有一個ESL評分樣版的課程成績 ， 不會有同一個ESL 評分樣版 有不同的課程成績
                    if (!_scoreDict[id].ContainsKey("課程學期成績分數"))
                    {
                        _scoreDict[id].Add("課程學期成績分數", score);
                    }
                    if (!_scoreDict[id].ContainsKey("課程學期成績等第"))
                    {
                        decimal score_d;
                        if (decimal.TryParse(score, out score_d))
                        {
                            _scoreDict[id].Add("課程學期成績等第", dm.GetDegreeByScore(score_d));
                        }
                    }

                    #endregion
                    
                }
            }



            #endregion


            foreach (K12.Data.SCAttendRecord scar in scList)
            {
                string id = scar.RefStudentID + "_" + scar.RefCourseID;

                string assessmentSetID = _courseIDPairDict[scar.RefCourseID];

                DataTable data = _assessmentSetupDataTableDict[assessmentSetID];
                
                DataRow row = data.NewRow();
                row["電子報表辨識編號"] = "系統編號{" + scar.Student.ID + "}"; // 學生系統編號

                row["學年度"] = scar.Course.SchoolYear;
                row["學期"] = scar.Course.Semester;
                row["學號"] = scar.Student.StudentNumber;
                row["年級"] = scar.Student.Class != null ? "" + scar.Student.Class.GradeYear : "";
                row["英文課程名稱"] = scar.Course.Name;
                row["原班級名稱"] = scar.Student.Class != null ? "" + scar.Student.Class.Name : "";
                row["學生英文姓名"] = scar.Student.EnglishName;
                row["學生中文姓名"] = scar.Student.Name;
                row["教師一"] = scar.Course.Teachers.Count > 0 ? scar.Course.Teachers.Find(x => x.Sequence == 1).TeacherName : ""; // 新寫法 直接找list 內教師條件
                row["教師二"] = scar.Course.Teachers.Count > 1 ? scar.Course.Teachers.Find(x => x.Sequence == 2).TeacherName : "";
                row["教師三"] = scar.Course.Teachers.Count > 2 ? scar.Course.Teachers.Find(x => x.Sequence == 3).TeacherName : "";

                if (_itemDict.ContainsKey(scar.RefCourseID))
                {
                    foreach (string mergeKey in _itemDict[scar.RefCourseID].Keys)
                    {
                        if (row.Table.Columns.Contains(mergeKey))
                        {
                            row[mergeKey] = _itemDict[scar.RefCourseID][mergeKey];
                        }                        
                    }
                }


                if (_scoreDict.ContainsKey(id))
                {
                    foreach (string mergeKey  in _scoreDict[id].Keys)
                    {
                        if (row.Table.Columns.Contains(mergeKey))
                        {
                            row[mergeKey] = _scoreDict[id][mergeKey];
                        }                        
                    }
                }

         

                data.Rows.Add(row);

                count++;
                progress += (int)(count * per);
                _worker.ReportProgress(progress);

            }

            Document docFinal = new Document();

            foreach (string assessmentSetupID in _assessmentSetupDataTableDict.Keys)
            {
                Document doc = _documentDict[assessmentSetupID];

                DataTable data = _assessmentSetupDataTableDict[assessmentSetupID];

                doc.MailMerge.Execute(data);

                docFinal.AppendDocument(doc, ImportFormatMode.KeepSourceFormatting);
            }


            docFinal.Sections[0].Remove();// 把第一頁刪掉


            e.Result = docFinal;


            _worker.ReportProgress(100, "ESL 報表列印完成。");
        }


        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage(" ESL成績單產生完成");

            Document doc = (Document)e.Result;
            doc.MailMerge.DeleteFields();

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
            sd.FileName = "ESL成績單.docx";
            sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    doc.Save(sd.FileName, Aspose.Words.SaveFormat.Docx);
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
            FISCA.Presentation.MotherForm.SetStatusBarMessage(""+e.UserState, e.ProgressPercentage);


        }



        // 產生 現在 所選課程 各自對應 ESL評分樣本設定 的功能變數
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

                    DataTable dataTable = new DataTable(); //要整理的資料，整理功能變數後要回傳

                    dataTable = GetMergeField(termList, _eslCouseList.FindAll(x => x.AssessmentSetup.ID == "" + dr["id"]));

                    if (!_assessmentSetupDataTableDict.ContainsKey("" + dr["id"]))
                    {
                        _assessmentSetupDataTableDict.Add("" + dr["id"], dataTable);
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


        private DataTable GetMergeField(List<Term> termList, List<K12.Data.CourseRecord> couseList)
        {
            // 計算權重用的字典(因為使用者在介面設定的權重數值 不一定就是想在報表上顯示的)
            // 目前康橋報表希望能夠將，每一個Subject、assessment 的比重 換算成為對於期中考的比例
            Dictionary<string, float> weightCalDict = new Dictionary<string, float>();

            DataTable dataTable = new DataTable();

            #region 固定變數
            // 固定變數，不分　期中、期末、學期
            // 基本資料
            dataTable.Columns.Add("學年度");
            dataTable.Columns.Add("學期");
            dataTable.Columns.Add("學號");
            dataTable.Columns.Add("年級");
            dataTable.Columns.Add("英文課程名稱");
            dataTable.Columns.Add("原班級名稱");
            dataTable.Columns.Add("學生英文姓名");
            dataTable.Columns.Add("學生中文姓名");
            dataTable.Columns.Add("教師一");
            dataTable.Columns.Add("教師二");
            dataTable.Columns.Add("教師三");
            dataTable.Columns.Add("電子報表辨識編號");
            #endregion


            // 學期課程成績
            if (!dataTable.Columns.Contains("課程學期成績分數"))
                dataTable.Columns.Add("課程學期成績分數");

            if (!dataTable.Columns.Contains("課程學期成績等第"))
                dataTable.Columns.Add("課程學期成績等第");

            // 2018/6/15 穎驊備註 以下整理 功能變數 最常使用的 string.Trim().Replace(' ', '_').Replace('"', '_') 
            // >> 其用意為避免Word 功能變數合併列印時 會有一些奇怪的BUG ，EX: row["Final-Term評量_Science科目_In-Class Score子項目_分數1"] = "YOYO!"; >> 有空格印不出來 

            foreach (K12.Data.CourseRecord course in couseList)
            {
                foreach (Term term in termList)
                {
                    if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重"))
                        dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重");

                    if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"))
                        dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"); // Term 分數本身 先暫時這樣處理之後要有類別整理

                    if (!_itemDict.ContainsKey(course.ID))
                    {
                        _itemDict.Add(course.ID, new Dictionary<string, string>());

                        _itemDict[course.ID].Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重", term.Weight);
                    }
                    else
                    {
                        _itemDict[course.ID].Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重", term.Weight);  //term 比重 
                    }

                    // 計算比重用，先整理 Term 、 Subject  的 總和
                    foreach (Subject subject in term.SubjectList)
                    {
                        // Term
                        if (!weightCalDict.ContainsKey(course.ID + "_" + term.Name + "_SubjectTotal"))
                        {
                            if (float.TryParse(subject.Weight, out float f))
                            {
                                weightCalDict.Add(course.ID + "_" + term.Name + "_SubjectTotal", f);
                            }
                        }
                        else
                        {
                            if (float.TryParse(subject.Weight, out float f))
                            {
                                weightCalDict[course.ID + "_" + term.Name + "_SubjectTotal"] += f;
                            }
                        }

                        // Subject
                        if (!weightCalDict.ContainsKey(course.ID + "_" + term.Name + "_" + subject.Name))
                        {
                            if (float.TryParse(subject.Weight, out float f))
                            {
                                weightCalDict.Add(course.ID + "_" + term.Name + "_" + subject.Name, f);
                            }
                        }

                    }

                    foreach (Subject subject in term.SubjectList)
                    {
                        if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重"))
                            dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重");
                        if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"))
                            dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"); // Subject 分數本身 先暫時這樣處理之後要有類別整理


                        string subjectWieght = "" + Math.Round((float.Parse(subject.Weight) * 100) / (weightCalDict[course.ID + "_" + term.Name + "_SubjectTotal"]), 2, MidpointRounding.ToEven);

                        _itemDict[course.ID].Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重", subjectWieght); //subject比重 

                        // 計算比重用，先整理 Assessment  的 總和
                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            if (!weightCalDict.ContainsKey(course.ID + "_" + term.Name + "_" + subject.Name + "_AssessmentTotal"))
                            {
                                if (float.TryParse(assessment.Weight, out float f))
                                {
                                    weightCalDict.Add(course.ID + "_" + term.Name + "_" + subject.Name + "_AssessmentTotal", f);
                                }
                            }
                            else
                            {
                                if (float.TryParse(assessment.Weight, out float f))
                                {
                                    weightCalDict[course.ID + "_" + term.Name + "_" + subject.Name + "_AssessmentTotal"] += f;
                                }
                            }
                        }



                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            if (assessment.Type == "Score") //分數型成績 才增加
                            {
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重");
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數");

                                string assessmentWieght = "" + Math.Round((weightCalDict[course.ID + "_" + term.Name + "_" + subject.Name] * float.Parse(assessment.Weight) * 100) / (weightCalDict[course.ID + "_" + term.Name + "_SubjectTotal"] * weightCalDict[course.ID + "_" + term.Name + "_" + subject.Name + "_AssessmentTotal"]), 2, MidpointRounding.ToEven);

                                _itemDict[course.ID].Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重", assessmentWieght); //assessment比重 

                            }
                            if (assessment.Type == "Indicator") // 檢查看有沒有　　Indicator　，有的話另外存List 做對照
                            {
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "指標"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "指標");

                                string key = course.ID + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_');

                                // 作為對照 key 為 courseID_termName_subjectName_assessment_Name
                                if (!_indicatorList.Contains(key))
                                {
                                    _indicatorList.Add(key);
                                }
                            }

                            if (assessment.Type == "Comment") // 檢查看有沒有　　Comment　，有的話另外存List 做對照
                            {
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "評語"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "評語");

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

            

            return dataTable;
        }


    }
}
