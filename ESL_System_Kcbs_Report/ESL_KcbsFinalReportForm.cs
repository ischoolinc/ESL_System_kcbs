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


namespace ESL_System_Kcbs_Report
{
    public partial class ESL_KcbsFinalReportForm : BaseForm
    {
        BackgroundWorker _BW;

        List<K12.Data.CourseRecord> esl_couse_list;

        // 儲放學生ESL 成績的dict 其結構為 <studentID_courseID,<scoreKey,scoreValue>
        Dictionary<string, Dictionary<string, string>> scoreDict = new Dictionary<string, Dictionary<string, string>>();

        public ESL_KcbsFinalReportForm(List<K12.Data.CourseRecord> _esl_couse_list)
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

                // 建立成績整理 Dict
                scoreDict.Add(scr.Student.ID + "_" + scr.Course.ID, new Dictionary<string, string>());
            }

            _BW.ReportProgress(20);


            int progress = 80;
            decimal per = (decimal)(100 - progress) / scList.Count;
            int count = 0;

            DataTable data = new DataTable();
            #region 整理合併欄位
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


            //Language Art
            //FET
            //Score
            data.Columns.Add("LA_FET_IN_S");
            data.Columns.Add("LA_FET_RAP_S");
            data.Columns.Add("LA_FET_MLAE_S");

            //Weight
            data.Columns.Add("LA_FET_IN_W");
            data.Columns.Add("LA_FET_RAP_W");
            data.Columns.Add("LA_FET_MLAE_W");

            //Performance
            data.Columns.Add("LA_FET_HC_P");
            data.Columns.Add("LA_FET_WWId_P");
            data.Columns.Add("LA_FET_WWIg_P");
            data.Columns.Add("LA_FET_B_P");

            //CET
            //Score
            data.Columns.Add("LA_CET_IN_S");
            data.Columns.Add("LA_CET_RAP_S");
            data.Columns.Add("LA_CET_MGE_S");

            //Weight
            data.Columns.Add("LA_CET_IN_W");
            data.Columns.Add("LA_CET_RAP_W");
            data.Columns.Add("LA_CET_MGE_W");

            //Performance
            data.Columns.Add("LA_CET_HC_P");
            data.Columns.Add("LA_CET_WWId_P");
            data.Columns.Add("LA_CET_WWIg_P");
            data.Columns.Add("LA_CET_B_P");

            //Final Score
            //data.Columns.Add("LA_F_S");
            data.Columns.Add("LA_F_W");

            //Science
            //FET
            //Score
            data.Columns.Add("SC_FET_IN_S");
            data.Columns.Add("SC_FET_SA_S");

            //Weight
            data.Columns.Add("SC_FET_IN_W");
            data.Columns.Add("SC_FET_SA_W");

            //Final Score
            //data.Columns.Add("SC_F_S");
            data.Columns.Add("SC_F_W");

            //Teacher’s Comments
            data.Columns.Add("LA_FET_C");
            data.Columns.Add("LA_CET_C");
            data.Columns.Add("SC_FET_C");

            //Total Score
            data.Columns.Add("M_W");
            data.Columns.Add("F_W");

            data.Columns.Add("LA_M_S");
            data.Columns.Add("SC_M_S");
            data.Columns.Add("SC_F_S");
            data.Columns.Add("LA_F_S");
            data.Columns.Add("LA_S_S");
            data.Columns.Add("SC_S_S");
            data.Columns.Add("LA_S_G");
            data.Columns.Add("SC_S_G");

            #endregion

            string course_ids = string.Join("','", courseIDList);

            string student_ids = string.Join("','", studentIDList);

            string sql = "SELECT * FROM $esl.gradebook_assessment_score WHERE ref_course_id IN ('" + course_ids + "') AND ref_student_id IN ('" + student_ids + "') AND term ='final-Term'";



            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(sql);
            foreach (DataRow row in dt.Rows)
            {
                string string1 = "";
                string string2 = "";

                string id = "" + row["ref_student_id"] + "_" + row["ref_course_id"];
                if (scoreDict.ContainsKey(id))
                {

                    switch ("" + row["subject"])
                    {
                        case "Language Art(FET)":
                            string1 = "LA_FET_";
                            break;
                        case "Language Art(CET)":
                            string1 = "LA_CET_";
                            break;
                        case "Science":
                            string1 = "SC_FET_";
                            break;
                        default:
                            break;
                    }

                    switch ("" + row["assessment"])
                    {
                        case "In-Class Score":
                            string2 = "IN_S";
                            break;
                        case "Reading Project":
                            string2 = "RAP_S";
                            break;
                        case "Language Arts Exam":
                            string2 = "MLAE_S";
                            break;
                        case "Grammar Exam":
                            string2 = "MGE_S";
                            break;
                        case "Project":
                            string2 = "SA_S";
                            break;
                        case "Comment":
                            string2 = "C";
                            break;
                        case "HomeWork Completion":
                            string2 = "HC_P";
                            break;
                        case "Works Well In a Group":
                            string2 = "WWIg_P";
                            break;
                        case "Works Well Independently":
                            string2 = "WWId_P";
                            break;
                        case "Behavior":
                            string2 = "B_P";
                            break;
                        default:
                            break;
                    }

                    if (!scoreDict[id].ContainsKey(string1 + string2))
                    {
                        scoreDict[id].Add(string1 + string2, "" + row["value"]);
                    }
                }

                //string id = "" + row["id"];
                //if (studentDic.ContainsKey(id))
                //{
                //    studentDic[id].MeritA = int.Parse("" + row["大功支數"]);
                //    studentDic[id].MeritB = int.Parse("" + row["小功支數"]);
                //    studentDic[id].MeritC = int.Parse("" + row["嘉獎支數"]);
                //    studentDic[id].DemeritA = int.Parse("" + row["大過支數"]);
                //    studentDic[id].DemeritB = int.Parse("" + row["小過支數"]);
                //    studentDic[id].DemeritC = int.Parse("" + row["警告支數"]);
                //}
            }


            foreach (K12.Data.SCAttendRecord scar in scList)
            {
                //百分比的問題統一在這處理
                SelectAssessmentSetup(scar);

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

                // Language  Art
                // FET Score
                row["LA_FET_IN_S"] = scoreDict[id].ContainsKey("LA_FET_IN_S") ? scoreDict[id]["LA_FET_IN_S"] : "";
                row["LA_FET_RAP_S"] = scoreDict[id].ContainsKey("LA_FET_RAP_S") ? scoreDict[id]["LA_FET_RAP_S"] : "";
                row["LA_FET_MLAE_S"] = scoreDict[id].ContainsKey("LA_FET_MLAE_S") ? scoreDict[id]["LA_FET_MLAE_S"] : "";

                // FET Weight
                row["LA_FET_IN_W"] = scoreDict[id].ContainsKey("LA_FET_IN_W") ? scoreDict[id]["LA_FET_IN_W"] + "%" : "";
                row["LA_FET_RAP_W"] = scoreDict[id].ContainsKey("LA_FET_RAP_W") ? scoreDict[id]["LA_FET_RAP_W"] + "%" : "";
                row["LA_FET_MLAE_W"] = scoreDict[id].ContainsKey("LA_FET_MLAE_W") ? scoreDict[id]["LA_FET_MLAE_W"] + "%" : "";


                // CET Score
                row["LA_CET_IN_S"] = scoreDict[id].ContainsKey("LA_CET_IN_S") ? scoreDict[id]["LA_CET_IN_S"] : "";
                row["LA_CET_RAP_S"] = scoreDict[id].ContainsKey("LA_CET_RAP_S") ? scoreDict[id]["LA_CET_RAP_S"] : "";
                row["LA_CET_MGE_S"] = scoreDict[id].ContainsKey("LA_CET_MGE_S") ? scoreDict[id]["LA_CET_MGE_S"] : "";

                // CET Weight
                row["LA_CET_IN_W"] = scoreDict[id].ContainsKey("LA_CET_IN_W") ? scoreDict[id]["LA_CET_IN_W"] + "%" : "";
                row["LA_CET_RAP_W"] = scoreDict[id].ContainsKey("LA_CET_RAP_W") ? scoreDict[id]["LA_CET_RAP_W"] + "%" : "";
                row["LA_CET_MGE_W"] = scoreDict[id].ContainsKey("LA_CET_MGE_W") ? scoreDict[id]["LA_CET_MGE_W"] + "%" : "";

                // Final Score
                row["LA_F_S"] = "";
                row["LA_F_W"] = "100%";

                // Science
                // FET Score
                row["SC_FET_IN_S"] = scoreDict[id].ContainsKey("SC_FET_IN_S") ? scoreDict[id]["SC_FET_IN_S"] : "";
                row["SC_FET_SA_S"] = scoreDict[id].ContainsKey("SC_FET_SA_S") ? scoreDict[id]["SC_FET_SA_S"] : "";

                // FET Weight
                row["SC_FET_IN_W"] = scoreDict[id].ContainsKey("SC_FET_IN_W") ? scoreDict[id]["SC_FET_IN_W"] +"%" : "";
                row["SC_FET_SA_W"] = scoreDict[id].ContainsKey("SC_FET_SA_W") ? scoreDict[id]["SC_FET_SA_W"] + "%" : "";

                // Final Score
                row["SC_F_S"] = "";
                row["SC_F_W"] = "100%";

                //Performance 
                row["LA_FET_HC_P"] = scoreDict[id].ContainsKey("LA_FET_HC_P") ? scoreDict[id]["LA_FET_HC_P"] : "";
                row["LA_FET_WWId_P"] = scoreDict[id].ContainsKey("LA_FET_WWId_P") ? scoreDict[id]["LA_FET_WWId_P"] : "";
                row["LA_FET_WWIg_P"] = scoreDict[id].ContainsKey("LA_FET_WWIg_P") ? scoreDict[id]["LA_FET_WWIg_P"] : "";
                row["LA_FET_B_P"] = scoreDict[id].ContainsKey("LA_FET_B_P") ? scoreDict[id]["LA_FET_B_P"] : "";

                row["LA_CET_HC_P"] = scoreDict[id].ContainsKey("LA_CET_HC_P") ? scoreDict[id]["LA_CET_HC_P"] : "";
                row["LA_CET_WWId_P"] = scoreDict[id].ContainsKey("LA_CET_WWId_P") ? scoreDict[id]["LA_CET_WWId_P"] : "";
                row["LA_CET_WWIg_P"] = scoreDict[id].ContainsKey("LA_CET_WWIg_P") ? scoreDict[id]["LA_CET_WWIg_P"] : "";
                row["LA_CET_B_P"] = scoreDict[id].ContainsKey("LA_CET_B_P") ? scoreDict[id]["LA_CET_B_P"] : "";


                //Comment
                row["LA_FET_C"] = scoreDict[id].ContainsKey("LA_FET_C") ? scoreDict[id]["LA_FET_C"] : "";
                row["LA_CET_C"] = scoreDict[id].ContainsKey("LA_CET_C") ? scoreDict[id]["LA_CET_C"] : "";
                row["SC_FET_C"] = scoreDict[id].ContainsKey("SC_FET_C") ? scoreDict[id]["SC_FET_C"] : "";



                //Semester conclusion
                row["M_W"] = scoreDict[id].ContainsKey("M_W") ? scoreDict[id]["M_W"] + "%" : "";
                row["F_W"] = scoreDict[id].ContainsKey("F_W") ? scoreDict[id]["F_W"] + "%" : "";                






                data.Rows.Add(row);

                count++;
                progress += (int)(count * per);
                _BW.ReportProgress(progress);
            }



            Document doc = new Document(new System.IO.MemoryStream(Properties.Resources.新_小學部英文成績單Report_Card_HC_Elementary_0317__期末_));
            doc.MailMerge.Execute(data);
            e.Result = doc;
        }

        private void ReportBuilding(object sender, RunWorkerCompletedEventArgs e)
        {
            MotherForm.SetStatusBarMessage(" ESL康橋期末成績單產生完成");

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
            sd.FileName = "ESL康橋期末成績單.docx";
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


        private void SelectAssessmentSetup(K12.Data.SCAttendRecord scar)
        {
            string xmlStr = "<root>" + scar.Course.AssessmentSetup.Description + "</root>";
            XElement elmRoot = XElement.Parse(xmlStr);

            string id = scar.RefStudentID + "_" + scar.RefCourseID;

            //解析讀下來的 descriptiony 資料，打包成物件群 最後交給 ParseDBxmlToNodeUI() 處理
            if (elmRoot != null)
            {
                if (elmRoot.Element("ESLTemplate") != null)
                {
                    foreach (XElement ele_term in elmRoot.Element("ESLTemplate").Elements("Term"))
                    {
                        if (ele_term.Attribute("Name").Value == "mid-Term")
                        {
                            if (!scoreDict[id].ContainsKey("M_W"))
                            {
                                scoreDict[id].Add("M_W", "" + ele_term.Attribute("Weight").Value );
                            }
                        }

                        if (ele_term.Attribute("Name").Value == "final-Term")
                        {
                            if (!scoreDict[id].ContainsKey("F_W"))
                            {
                                scoreDict[id].Add("F_W", "" + ele_term.Attribute("Weight").Value );
                            }


                            foreach (XElement ele_subject in ele_term.Elements("Subject"))
                            {
                                
                                string string1 = "";


                                switch ("" + ele_subject.Attribute("Name").Value)
                                {
                                    case "Language Art(FET)":
                                        string1 = "LA_FET_";
                                        break;
                                    case "Language Art(CET)":
                                        string1 = "LA_CET_";
                                        break;
                                    case "Science":
                                        string1 = "SC_FET_";
                                        break;
                                    default:
                                        break;
                                }

                                foreach (XElement ele_assessment in ele_subject.Elements("Assessment"))
                                {
                                    string string2 = "";

                                    switch ("" + ele_assessment.Attribute("Name").Value)
                                    {
                                        case "In-Class Score":
                                            string2 = "IN_W";
                                            break;
                                        case "Reading Project":
                                            string2 = "RAP_W";
                                            break;
                                        case "Language Arts Exam":
                                            string2 = "MLAE_W";
                                            break;
                                        case "Grammar Exam":
                                            string2 = "MGE_W";
                                            break;
                                        case "Project":
                                            string2 = "SA_W";
                                            break;                                  
                                        default:
                                            break;
                                    }

                                    if (!scoreDict[id].ContainsKey(string1+string2))
                                    {
                                        scoreDict[id].Add(string1+string2, "" + ele_assessment.Attribute("Weight").Value );
                                    }

                                }
                               
                            }
                        }
                    }
 
                }
            }

        }

    }
}
