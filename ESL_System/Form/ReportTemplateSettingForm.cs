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
using Aspose.Words;
using System.IO;
using System.Xml.Linq;

namespace ESL_System.Form
{
    public partial class ReportTemplateSettingForm : BaseForm
    {
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

        public string SourceID { set; get; }

        public string SourceType { set; get; } //分成 期中、期末、學期 三種

        private List<UDT_ReportTemplate> _Configures = new List<UDT_ReportTemplate>();
        public UDT_ReportTemplate Configure { get; set; }

        public ReportTemplateSettingForm()
        {
            InitializeComponent();

            // 因應康橋國小學制(6年)，提供目前學年度往後六年的學年度選擇
            comboBoxEx1.Items.Add(int.Parse(School.DefaultSchoolYear) - 6);
            comboBoxEx1.Items.Add(int.Parse(School.DefaultSchoolYear) - 5);
            comboBoxEx1.Items.Add(int.Parse(School.DefaultSchoolYear) - 4);
            comboBoxEx1.Items.Add(int.Parse(School.DefaultSchoolYear) - 3);
            comboBoxEx1.Items.Add(int.Parse(School.DefaultSchoolYear) - 2);
            comboBoxEx1.Items.Add(int.Parse(School.DefaultSchoolYear) - 1);
            comboBoxEx1.Items.Add(int.Parse(School.DefaultSchoolYear));

            comboBoxEx2.Items.Add(1);
            comboBoxEx2.Items.Add(2);

            // 預設為學校的當學年度學期
            comboBoxEx1.Text = School.DefaultSchoolYear;
            comboBoxEx2.Text = School.DefaultSemester;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ReportTemplateSettingForm_Load(object sender, EventArgs e)
        {
            if (SourceType == "期中")
            {
                this.Text = "期中報表設定";
            }
            if (SourceType == "期末")
            {
                this.Text = "期末報表設定";
            }
            if (SourceType == "學期")
            {
                this.Text = "學期報表設定";
            }

        }

        // 依照所選的 評分樣板 、報表種類 、 學年度、學期 Load 進正確的 Configure
        private void LoadConitionalConfigure()
        {
            string qry = "ref_exam_template_id ='" + SourceID + "' and schoolyear='" + comboBoxEx1.Text + "' and semester ='" + comboBoxEx2.Text + "' and exam ='" + SourceType + "'";
            _Configures = _AccessHelper.Select<UDT_ReportTemplate>(qry);

            Configure = _Configures.Count > 0 ? _Configures[0] : null; // 理論上只會有一筆
        }

        // 檢視套印樣板
        private void lbv01_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LoadConitionalConfigure();

            // 假如沒有抓到 Configure 代表是第一次設定，帶入預設的報表後，也幫她上傳儲存設定
            if (Configure == null)
            {
                Configure = new UDT_ReportTemplate(); // 沒有的話 就帶一個新的給它

                Configure.Ref_exam_Template_ID = SourceID;
                Configure.SchoolYear = comboBoxEx1.Text;
                Configure.Semester = comboBoxEx2.Text;
                Configure.Exam = SourceType;

                if (SourceType == "期中")
                {
                    //Configure.Template = new Document(new System.IO.MemoryStream(Properties.Resources.MidReport));
                    // 2018/10/15 穎驊註記，不再帶出預設定的樣板，因為每間學校不一樣，跑出別的設定很奇怪， 
                    // 沒設定的話 就給她空的
                    Configure.Template = new Document();
                }
                if (SourceType == "期末")
                {
                    //Configure.Template = new Document(new System.IO.MemoryStream(Properties.Resources.FinalReport));
                    // 2018/10/15 穎驊註記，不再帶出預設定的樣板，因為每間學校不一樣，跑出別的設定很奇怪， 
                    // 沒設定的話 就給她空的
                    Configure.Template = new Document();
                }
                if (SourceType == "學期")
                {
                    //Configure.Template = new Document(new System.IO.MemoryStream(Properties.Resources.SemesterReport));
                    // 2018/10/15 穎驊註記，不再帶出預設定的樣板，因為每間學校不一樣，跑出別的設定很奇怪， 
                    // 沒設定的話 就給她空的
                    Configure.Template = new Document();
                }

                Configure.Encode(); // 將Word 轉成 stream

                Configure.Save();
            }

            Document doc = new Document();

            Configure.Decode(); // 將 stream 轉成 Word

            doc = Configure.Template;

            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = SourceType + "報表樣版.docx";
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

        // 變更套印樣板
        private void lbc01_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LoadConitionalConfigure();

            // 假如沒有抓到 Configure 代表是第一次設定，帶入預設的報表後，也幫她上傳儲存設定
            if (Configure == null)
            {
                Configure = new UDT_ReportTemplate(); // 沒有的話 就帶一個新的給它

                Configure.Ref_exam_Template_ID = SourceID;
                Configure.SchoolYear = comboBoxEx1.Text;
                Configure.Semester = comboBoxEx2.Text;
                Configure.Exam = SourceType;

                if (SourceType == "期中")
                {
                    Configure.Template = new Document(new System.IO.MemoryStream(Properties.Resources.MidReport));
                }
                if (SourceType == "期末")
                {
                    Configure.Template = new Document(new System.IO.MemoryStream(Properties.Resources.FinalReport));
                }
                if (SourceType == "學期")
                {
                    Configure.Template = new Document(new System.IO.MemoryStream(Properties.Resources.SemesterReport));
                }

                Configure.Encode(); // 將Word 轉成 stream

                Configure.Save();
            }

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    Configure.Template = new Aspose.Words.Document(dialog.FileName);

                    Configure.Encode(); // 將word 轉成 stream
                    Configure.Save(); // 儲存

                    MsgBox.Show("樣板更改成功!");
                }
                catch
                {
                    MsgBox.Show("樣板開啟失敗!");
                }
            }


        }


        // 檢視功能變數(按照不同的SourceType 回傳不同的 功能變數樣版)
        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            #region 舊寫法　功能變數完全寫死在專案裡
            //Document doc = new Document();

            //if (SourceType == "期中")
            //{
            //    doc = new Document(new System.IO.MemoryStream(Properties.Resources.MidReport_MailMergeList));
            //}
            //if (SourceType == "期末")
            //{
            //    doc = new Document(new System.IO.MemoryStream(Properties.Resources.FinalReport_MailMergeList));
            //}
            //if (SourceType == "學期")
            //{
            //    doc = new Document(new System.IO.MemoryStream(Properties.Resources.SemesterReport_MailMergeList));
            //}

            //SaveFileDialog sd = new SaveFileDialog();
            //sd.Title = "另存新檔";
            //sd.FileName = SourceType + "報表功能變數總表.docx";
            //sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            //if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    try
            //    {
            //        doc.Save(sd.FileName, Aspose.Words.SaveFormat.Docx);
            //        System.Diagnostics.Process.Start(sd.FileName);
            //    }
            //    catch
            //    {
            //        MessageBox.Show("檔案儲存失敗");
            //    }
            //} 
            #endregion

            // 新寫法，依照　所選取的ESL 樣板設定，產生出動態對應階層的功能變數總表！！
            CreateFieldTemplate();
            return;

        }


        // 穎驊搬過來的工具，可以一次大量建立有規則的功能變數，可以省下很多時間。
        private void CreateFieldTemplate()
        {

            List<Term> termList = new List<Term>();

            #region  解讀　description　XML
            // 取得ESL 描述 in description
            DataTable dt;
            QueryHelper qh = new QueryHelper();

            string selQuery = "select id,description from exam_template where id = '" + SourceID + "'";
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
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);

            #region 固定變數
            // 固定變數，不分　期中、期末、學期
            builder.Write("固定變數");
            builder.StartTable();
            builder.InsertCell();
            builder.Write("項目");
            builder.InsertCell();
            builder.Write("變數");
            builder.EndRow();
            foreach (string key in new string[]{
                    "學年度",
                    "學期",
                    "學號",
                    "年級",
                    "英文課程名稱",
                    "原班級名稱",
                    "學生英文姓名",
                    "學生中文姓名",
                    "教師一",
                    "教師二",
                    "教師三",
                    "電子報表辨識編號"
                })
            {
                builder.InsertCell();
                builder.Write(key);
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + key + " \\* MERGEFORMAT ", "«" + key + "»");
                builder.EndRow();
            }

            builder.EndTable();

            builder.Writeln();
            #endregion

            #region 成績變數

            int termCounter = 1;

            // 2018/6/15 穎驊備註 以下整理 功能變數 最常使用的 string..Trim().Replace(' ', '_').Replace('"', '_') 
            // >> 其用意為避免Word 功能變數合併列印時 會有一些奇怪的BUG ，EX: row["Final-Term評量_Science科目_In-Class Score子項目_分數1"] = "YOYO!"; >> 有空格印不出來 

            foreach (Term term in termList)
            {                
                builder.Writeln(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "評量");

                builder.StartTable();
                builder.InsertCell();
                builder.Write("評量名稱");
                builder.InsertCell();
                builder.Write("評量分數");
                builder.InsertCell();
                builder.Write("評量比重");
                builder.EndRow();

                builder.InsertCell();
                //builder.InsertField("MERGEFIELD " + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "名稱" + termCounter + " \\* MERGEFORMAT ", "«I" + termCounter + "»");
                builder.Write(term.Name);

                builder.InsertCell();
                //builder.InsertField("MERGEFIELD " + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "分數" + termCounter + " \\* MERGEFORMAT ", "«TS" + termCounter + "»");
                builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數" + " \\* MERGEFORMAT ", "«TS»");

                builder.InsertCell();
                //builder.InsertField("MERGEFIELD " + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "比重" + termCounter+ " \\* MERGEFORMAT ", "«TW" + termCounter + "»");
                builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數" + " \\* MERGEFORMAT ", "«TW»");

                //termCounter++;

                builder.EndRow();
                builder.EndTable();

                builder.Writeln();

                int subjectCounter = 1;

                foreach (Subject subject in term.SubjectList)
                {
                    builder.Writeln(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "科目分數型成績");

                    builder.StartTable();
                    builder.InsertCell();
                    builder.Write("科目名稱");
                    builder.InsertCell();
                    builder.Write("科目分數");
                    builder.InsertCell();
                    builder.Write("科目比重");                    
                    builder.EndRow();


                    builder.InsertCell();
                    builder.Write(subject.Name);
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數" + " \\* MERGEFORMAT ", "«SS»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重" + " \\* MERGEFORMAT ", "«SW»");

                    //subjectCounter++;

                    builder.EndRow();
                    builder.EndTable();

                    builder.StartTable();
                    builder.InsertCell();
                    builder.Write("子項目名稱");
                    builder.InsertCell();
                    builder.Write("比重");
                    builder.InsertCell();
                    builder.Write("分數");

                    builder.EndRow();

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

                        if (assessment.Type != "Score") //  非分數型成績 跳過 不寫入
                        {
                            continue;
                        }


                        builder.InsertCell();
                        //builder.InsertField("MERGEFIELD " + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "名稱" + assessmentCounter + " \\* MERGEFORMAT ", "«I" + assessmentCounter + "»");
                        builder.Write(assessment.Name);

                        builder.InsertCell();
                        //builder.InsertField("MERGEFIELD " + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "比重" + assessmentCounter + " \\* MERGEFORMAT ", "«AW" + assessmentCounter + "»");
                        builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重"  + " \\* MERGEFORMAT ", "«AW»");

                        builder.InsertCell();
                        //builder.InsertField("MERGEFIELD " + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "分數" + assessmentCounter + " \\* MERGEFORMAT ", "«S" + assessmentCounter + "»");
                        builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"  + " \\* MERGEFORMAT ", "«AS»");

                        assessmentCounter++;

                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln();


                    // 處理Indicator
                    if (assessmentContainsIndicator)
                    {
                        builder.Writeln(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "指標型成績");

                        builder.StartTable();
                        builder.InsertCell();
                        builder.Write("項目");
                        builder.InsertCell();
                        builder.Write("指標");
                        builder.EndRow();

                        assessmentCounter = 1;
                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            if (assessment.Type == "Indicator") // 檢查看有沒有　Indicator　，專為 Indicator 畫張表
                            {
                                builder.InsertCell();
                                builder.Write(assessment.Name);
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "指標" + " \\* MERGEFORMAT ", "«I»");
                                builder.EndRow();
                                //assessmentCounter++;
                            }
                        }
                        builder.EndTable();
                        builder.Writeln();
                    }

                    // 處理Comment
                    if (assessmentContainsComment)
                    {
                        builder.Writeln(term.Name + "/" + subject.Name + "評語型成績");

                        builder.StartTable();
                        builder.InsertCell();
                        builder.Write("項目");
                        builder.InsertCell();
                        builder.Write("評語");
                        builder.EndRow();

                        assessmentCounter = 1;
                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            if (assessment.Type == "Comment") // 檢查看有沒有　Comment　，專為 Comment 畫張表
                            {
                                builder.InsertCell();
                                builder.Write(assessment.Name);
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "評語" +" \\* MERGEFORMAT ", "«C»");
                                builder.EndRow();
                                assessmentCounter++;
                            }
                        }
                        builder.EndTable();
                        builder.Writeln();
                    }

                }
                builder.Writeln();
            }

            #endregion

            #region 課程學期成績
            builder.Writeln("課程學期成績");

            builder.StartTable();
            builder.InsertCell();
            builder.Write("課程學期成績分數");
            builder.InsertCell();
            builder.Write("課程學期成績等第");
            builder.EndRow();

            builder.InsertCell();

            builder.InsertField("MERGEFIELD "  + "課程學期成績分數" + " \\* MERGEFORMAT ", "«CSS»");

            builder.InsertCell();

            builder.InsertField("MERGEFIELD "  + "課程學期成績等第" + " \\* MERGEFORMAT ", "«CSL»");

            builder.EndRow();
            builder.EndTable();
            #endregion


            #region 儲存檔案
            string inputReportName = "合併欄位總表";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

            if (System.IO.File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!System.IO.File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                doc.Save(path, Aspose.Words.SaveFormat.Doc);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Excel檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        doc.Save(path, Aspose.Words.SaveFormat.Doc);
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            #endregion
        }

    }
}

