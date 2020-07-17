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
using FISCA.Presentation;
using K12.Data;
using K12.Presentation;
using Aspose.Words;
using System.Xml.Linq;


namespace ESL_System
{
    public partial class ESL_KcbsFinalReportFormNEW : BaseForm
    {

        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

        BackgroundWorker _BW;

        List<K12.Data.CourseRecord> esl_couse_list;

        // 儲放學生ESL 成績的dict 其結構為 <studentID_courseID,<scoreKey,scoreValue>
        Dictionary<string, Dictionary<string, string>> scoreDict = new Dictionary<string, Dictionary<string, string>>();

        // 儲放ESL 成績單 科目、比重設定 的dict 其結構為 <key,value>
        Dictionary<string, string> itemDict = new Dictionary<string, string>();


        Dictionary<string, int> scoreItemSortingDict = new Dictionary<string, int>(); //將樣板的 成績項目做出排序 如:<"Mid-Term",1>,<"Final-Term",2>

        //Dictionary<string, List<string>> courseIndicatorDict = new Dictionary<string, List<string>>();  // 指出該課程有哪些 Assessment 屬於 Indicator 成績

        List<string> IndicatorList = new List<string>();// 暫時用來判斷而者為Indicator 項目 之後需要重新設計 批次的做法

        

        public ESL_KcbsFinalReportFormNEW(List<K12.Data.CourseRecord> _esl_couse_list)
        {
            InitializeComponent();

            esl_couse_list = _esl_couse_list;

            _BW = new BackgroundWorker();
            _BW.WorkerReportsProgress = true;
            _BW.DoWork += new DoWorkEventHandler(DataBuilding);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ReportBuilding);
            _BW.ProgressChanged += new ProgressChangedEventHandler(BW_Progress);

            _BW.RunWorkerAsync();
        }

        private void BW_Progress(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("ESL康橋期末成績單產生完成產生中", e.ProgressPercentage);
        }


        private void DataBuilding(object sender, DoWorkEventArgs e)
        {
            string qry = "ref_exam_template_id ='" + esl_couse_list[0].RefAssessmentSetupID + "' and schoolyear='" + K12.Data.School.DefaultSchoolYear + "' and semester ='" + K12.Data.School.DefaultSemester + "' and exam ='期末'";

            List<UDT_ReportTemplate> _Configures = _AccessHelper.Select<UDT_ReportTemplate>(qry);

            UDT_ReportTemplate Configure;
            Document doc = new Document();

            Configure = _Configures.Count > 0 ? _Configures[0] : null; // 理論上只會有一筆

            Configure.Decode(); // 將 stream 轉成 Word

            doc = Configure.Template;


            _BW.ReportProgress(0);

            List<string> courseIDList = new List<string>();
            List<string> studentIDList = new List<string>();


            foreach (K12.Data.CourseRecord cr in esl_couse_list)
            {
                courseIDList.Add(cr.ID);
            }

            List<K12.Data.SCAttendRecord> scList = K12.Data.SCAttend.SelectByCourseIDs(courseIDList);


            foreach (K12.Data.SCAttendRecord scr in scList)
            {
                studentIDList.Add(scr.Student.ID);

                // 建立成績整理 Dict ，[studentID_courseID,[scoreKey,scoreID]]
                scoreDict.Add(scr.Student.ID + "_" + scr.Course.ID, new Dictionary<string, string>());
            }

            _BW.ReportProgress(20);


            int progress = 80;
            decimal per = (decimal)(100 - progress) / scList.Count;
            int count = 0;

            DataTable data = CreateFieldTemplate();
                        
            string course_ids = string.Join("','", courseIDList);

            string student_ids = string.Join("','", studentIDList);

            //string sql = "SELECT * FROM $esl.gradebook_assessment_score WHERE ref_course_id IN ('" + course_ids + "') AND ref_student_id IN ('" + student_ids + "') AND term ='Final-Term'"; // 這邊先暫訂試別 為Final-Term ， 之後要動態抓取

            string sql = "SELECT * FROM $esl.gradebook_assessment_score WHERE ref_course_id IN ('" + course_ids + "') AND ref_student_id IN ('" + student_ids + "') "; // 2018/6/21 通通都抓了，因為一張成績單上資訊，不只Final的

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(sql);

            int assessmentCounter = 1;

            foreach (DataRow row in dt.Rows)
            {
                string indexKey = "";

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
                    indexKey = termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_');

                    if (scoreDict.ContainsKey(id))
                    {

                        if (IndicatorList.Contains(assessmentWord)) // Indicator 子項目成績暫時的處理方式
                        {
                            scoreDict[id].Add(termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_') + "子項目_指標" + (IndicatorList.IndexOf(assessmentWord) + 1), "" + row["value"]);

                            assessmentCounter++;
                        }
                        else if (assessmentWord == "Comments") // 評語的暫時處理方式
                        {
                            scoreDict[id].Add(termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_') + "子項目_評語1" , "" + row["value"]);

                            assessmentCounter++;
                        }
                        else
                        {
                            scoreDict[id].Add(termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_') + "分數" + scoreItemSortingDict[indexKey], "" + row["value"]);

                            assessmentCounter++;
                        }

                        
                    }

                }

                // 沒有assessment，為subject 成績
                if (termWord != "" && "" + subjectWord != "" && "" + assessmentWord == "")
                {
                    indexKey = termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_');

                    if (scoreDict.ContainsKey(id))
                    {
                        scoreDict[id].Add(termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "分數" + scoreItemSortingDict[indexKey], "" + row["value"]);

                        assessmentCounter++;
                    }
                }
                // 沒有assessment、subject，為term 成績
                if (termWord != "" && "" + subjectWord == "" && "" + assessmentWord == "")
                {
                    indexKey = termWord.Trim().Replace(' ', '_').Replace('"', '_');

                    if (scoreDict.ContainsKey(id))
                    {
                        scoreDict[id].Add(termWord.Trim().Replace(' ', '_').Replace('"', '_') + "分數" + scoreItemSortingDict[indexKey], "" + row["value"]);

                        assessmentCounter++;
                    }
                }
         
                

            }


            foreach (K12.Data.SCAttendRecord scar in scList)
            {
                //百分比的問題統一在這處理
                //SelectAssessmentSetup(scar);

                string id = scar.RefStudentID + "_" + scar.RefCourseID;

                DataRow row = data.NewRow();
                //row["電子報表辨識編號"] = "系統編號{" + scar.Student.ID + "}"; // 學生系統編號

                row["學號"] = scar.Student.StudentNumber;
                row["年級"] = scar.Student.Class != null ? "" + scar.Student.Class.GradeYear : "";
                row["英文課程名稱"] = scar.Course.Name;
                row["原班級名稱"] = scar.Student.Class != null ? "" + scar.Student.Class.Name : "";
                row["學生英文姓名"] = scar.Student.EnglishName;
                row["學生中文姓名"] = scar.Student.Name;
                row["教師一"] = scar.Course.Teachers.Count > 0 ? scar.Course.Teachers.Find(x => x.Sequence == 1).TeacherName : ""; // 新寫法 直接找list 內教師條件
                row["教師二"] = scar.Course.Teachers.Count > 1 ? scar.Course.Teachers.Find(x => x.Sequence == 2).TeacherName : "";
                row["教師三"] = scar.Course.Teachers.Count > 2 ? scar.Course.Teachers.Find(x => x.Sequence == 3).TeacherName : "";


                foreach (KeyValuePair<string, string> p in itemDict)
                {
                    row[p.Key] = p.Value;
                }

                foreach (KeyValuePair<string, string> p in scoreDict[id])
                {
                    row[p.Key] = p.Value;
                }

                


                //row["123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890"] = "YOYO1";
                //row["Final-Term評量_Science科目_Project子項目_分數2"] = "YOYO";

                //row["Final-Term評量_Science科目_In-ClassScore子項目_分數1"] = "YOYO!";

                //row["Final-Term評量_Science科目_Project子項目_分數2"] = "YOYO";


                data.Rows.Add(row);

                count++;
                progress += (int)(count * per);
                _BW.ReportProgress(progress);
            }

            // 這邊還是先用 Local 的，因為設定樣板 沒有選擇試別 不好定位
            //Document doc = new Document(new System.IO.MemoryStream(Properties.Resources.FinalReport));
            //Document doc = new Document();
            doc.MailMerge.Execute(data);
            e.Result = doc;
        }

        // 2018/6/14 目前只能選單筆的ESL 課程列印，之後要改
        // 穎驊搬過來的工具，可以一次大量建立有規則的功能變數，可以省下很多時間。
        private DataTable CreateFieldTemplate()
        {
            DataTable data = new DataTable(); //要整理的資料，整理功能變數後要回傳

            List<Term> termList = new List<Term>();

            #region  解讀　description　XML
            // 取得ESL 描述 in description
            DataTable dt;
            QueryHelper qh = new QueryHelper();

            string selQuery = "select id,description from exam_template where id = '" + esl_couse_list[0].RefAssessmentSetupID + "'";
            dt = qh.Select(selQuery);
            string xmlStr = "<root>" + dt.Rows[0]["description"].ToString() + "</root>";
            XElement elmRoot = XElement.Parse(xmlStr);

            //解析讀下來的 descriptiony 資料，打包成物件群 最後交給 ParseDBxmlToNodeUI() 處理
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

                        termList.Add(t); // 整理成大包的termList 後面用此來拚功能變數總表
                    }
                }
            }
            #endregion

            Aspose.Words.Document doc = new Aspose.Words.Document();

            

            #region 固定變數
            // 固定變數，不分　期中、期末、學期
            // 基本資料
            data.Columns.Add("學號");
            data.Columns.Add("年級");
            data.Columns.Add("英文課程名稱");
            data.Columns.Add("原班級名稱");
            data.Columns.Add("學生英文姓名");
            data.Columns.Add("學生中文姓名");
            data.Columns.Add("教師一");
            data.Columns.Add("教師二");
            data.Columns.Add("教師三");
            #endregion

            #region 成績變數

            int termCounter = 1;

            // 2018/6/15 穎驊備註 以下整理 功能變數 最常使用的 string.Trim().Replace(' ', '_').Replace('"', '_') 
            // >> 其用意為避免Word 功能變數合併列印時 會有一些奇怪的BUG ，EX: row["Final-Term評量_Science科目_In-Class Score子項目_分數1"] = "YOYO!"; >> 有空格印不出來 
            foreach (Term term in termList)
            {
                data.Columns.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "名稱" + termCounter);
                data.Columns.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "比重" + termCounter);
                data.Columns.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "分數" + termCounter); // Term 分數本身 先暫時這樣處理之後要有類別整理

                itemDict.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "名稱" + termCounter, term.Name.Trim().Replace(' ', '_').Replace('"', '_')); //Term 名稱 
                itemDict.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "比重" + termCounter, term.Weight); //Term 比重 

                scoreItemSortingDict.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_'), termCounter); //成績子項目 順位的整理
                
                termCounter++;

                int subjectCounter = 1;

                foreach (Subject subject in term.SubjectList)
                {
                    data.Columns.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "名稱" + subjectCounter);
                    data.Columns.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_')  + "比重" + subjectCounter);
                    data.Columns.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "分數" + subjectCounter); // Subject 分數本身 先暫時這樣處理之後要有類別整理

                    itemDict.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "名稱" + subjectCounter, subject.Name.Trim().Replace(' ', '_').Replace('"', '_')); //subject名稱 
                    itemDict.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "比重" + subjectCounter, subject.Weight); //subject比重 

                    scoreItemSortingDict.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_'), subjectCounter); //成績子項目 順位的整理

                    subjectCounter++;

                    int assessmentCounter = 1;

                    bool assessmentContainsIndicator = false;

                    bool assessmentContainsComment = false;

                    foreach (Assessment assessment in subject.AssessmentList)
                    {
                        if (assessment.Type == "Indicator") // 檢查看有沒有　　Indicator　，有的話，會另外再畫一張表專放Indicator                       
                        {
                            assessmentContainsIndicator = true;
                        }

                        if (assessment.Type == "Comment") // 檢查看有沒有　　Comment　，有的話，會另外再畫一張表專放Comment                        
                        {
                            assessmentContainsComment = true;
                        }

                        data.Columns.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "名稱" + assessmentCounter);
                        data.Columns.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "比重" + assessmentCounter);
                        data.Columns.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "分數" + assessmentCounter);

                        itemDict.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "名稱" + assessmentCounter, assessment.Name.Trim().Replace(' ', '_').Replace('"', '_')); //assessment名稱 
                        itemDict.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "比重" + assessmentCounter, assessment.Weight); //assessment比重

                        scoreItemSortingDict.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_'), assessmentCounter); //成績子項目 順位的整理

                        assessmentCounter++;
                        
                    }
    
                    // 處理Indicator
                    if (assessmentContainsIndicator)
                    {                                                                       
                        assessmentCounter = 1;
                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            if (assessment.Type == "Indicator") // 檢查看有沒有　Indicator　，專為 Indicator 畫張表
                            {
                                data.Columns.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "子項目_指標" + assessmentCounter);

                                if (!IndicatorList.Contains(assessment.Name)) // 暫且的整理
                                {
                                    IndicatorList.Add(assessment.Name);
                                }
                                
                                assessmentCounter++;
                            }
                        }                        
                    }

                    // 處理Comment
                    if (assessmentContainsComment)
                    {
                        assessmentCounter = 1;
                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            if (assessment.Type == "Comment") // 檢查看有沒有　Comment　，專為 Comment 畫張表
                            {
                                data.Columns.Add(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "子項目_評語" + assessmentCounter);                                

                                assessmentCounter++;
                            }
                        }                        
                    }

                }                
            }

            #endregion
            return data;
        }


        private void ReportBuilding(object sender, RunWorkerCompletedEventArgs e)
        {
            MotherForm.SetStatusBarMessage(" ESL成績單產生完成");

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
        }


        

    }
}
