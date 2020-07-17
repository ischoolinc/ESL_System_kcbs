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
using DevComponents.DotNetBar;
using System.Xml.Linq;
using K12.Data;
using System.Xml;
using System.IO;
using Aspose.Words;
using K12.Data.Configuration;
using ESL_System.Service;
using ESL_System.Model;
using FISCA.UDT;
using ESL_System.UDT;

namespace ESL_System.Form
{
    public partial class PrintStudentESLReportForm : FISCA.Presentation.Controls.BaseForm
    {
        private BackgroundWorker _bw = new BackgroundWorker();

        private string _wordURL = "";

        private bool _isOrderBySubject = false;
        private string _orderedSubject = ""; // 排序科目

        private string school_year = "";
        private string semester = "";

        // 儲放 學期科目成績算術平均用物件
        private SemsTotalScoreInfos _SemsTotalScoreInfos = new SemsTotalScoreInfos();

        // 儲放學生ESL 成績的dict 其結構為 <studentID,<scoreKey,scoreValue>
        private Dictionary<string, Dictionary<string, string>> _scoreDict = new Dictionary<string, Dictionary<string, string>>();

        // 儲放ESL 成績單 科目、比重設定 的dict 
        private Dictionary<string, Dictionary<string, string>> _itemDict = new Dictionary<string, Dictionary<string, string>>();

        // 紀錄成績 為 指標型indicator 的 key值 ， 作為對照 key 為 courseID_termName_subjectName_assessment_Name
        private List<string> _indicatorList = new List<string>();

        // 紀錄成績 為 評語型comment 的 key值
        private List<string> _commentList = new List<string>();

        // 儲放教務作業系統設定的 評量名稱
        private List<string> _examList = new List<string>();

        // 儲放教務作業系統設定的 領域名稱
        private List<string> _doaminList = new List<string>();

        // 儲放 所有課程的科目名稱
        private List<string> _subjectList = new List<string>();

        // 缺曠區間統計
        Dictionary<string, Dictionary<string, int>> _AttendanceDict = new Dictionary<string, Dictionary<string, int>>();

        // 獎懲統計
        Dictionary<string, Dictionary<string, int>> _DisciplineCountDict = new Dictionary<string, Dictionary<string, int>>();

        BackgroundWorker bkw;


        // 開始日期
        private DateTime _BeginDate;
        // 結束日期
        private DateTime _EndDate;


        private List<string> _typeList = new List<string>();
        private List<string> _absenceList = new List<string>();


        private Document _doc;

        private DataTable _mergeDataTable = new DataTable();

        private List<UDT_ReportTemplate> _configuresList = new List<UDT_ReportTemplate>();

        private UDT_ReportTemplate _configure { get; set; }


        public PrintStudentESLReportForm()
        {
            InitializeComponent();

            _bw.DoWork += new DoWorkEventHandler(_bkWork_DoWork);
            _bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_worker_RunWorkerCompleted);
            _bw.ProgressChanged += new ProgressChangedEventHandler(_worker_ProgressChanged);
            _bw.WorkerReportsProgress = true;

            bkw = new BackgroundWorker();
            bkw.DoWork += new DoWorkEventHandler(bkw_DoWork);
            bkw.ProgressChanged += new ProgressChangedEventHandler(bkw_ProgressChanged);
            bkw.WorkerReportsProgress = true;
            bkw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bkw_RunWorkerCompleted);

            // 預設都為今天
            dtBegin.Value = DateTime.Now;
            dtEnd.Value = DateTime.Now;


            // 缺曠資料
            foreach (PeriodMappingInfo info in PeriodMapping.SelectAll())
            {
                if (!_typeList.Contains(info.Type))
                    _typeList.Add(info.Type);
            }

            foreach (AbsenceMappingInfo info in AbsenceMapping.SelectAll())
            {
                if (!_absenceList.Contains(info.Name))
                    _absenceList.Add(info.Name);
            }

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

            btnPrint.Enabled = false;
        }

        void bkw_DoWork(object sender, DoWorkEventArgs e)
        {
            bkw.ReportProgress(1);

            QueryHelper qh = new QueryHelper();


            #region 抓科目
            // 選出 所有有設定ESL 評分樣版 的科目名稱 

            bkw.ReportProgress(20);
            string sqlCourseID = @"
    SELECT 		
	DISTINCT	course.subject
	FROM course 
	WHERE 	
     ref_exam_template_id IN(
		SELECT id 
		FROM exam_template  
		WHERE description IS NOT NULL)
	AND course.subject IS NOT NULL";

            DataTable dtCourseID = qh.Select(sqlCourseID);

            foreach (DataRow row in dtCourseID.Rows)
            {
                string subject = "" + row["subject"];
                _subjectList.Add(subject);
            }
            #endregion

            #region 抓試別
            // 2018/12/18 穎驊註解，為了提供序列式的功能變數抓出系統，
            // 這邊填入 教務作業設定的 所有評量名稱

            bkw.ReportProgress(40);
            string sqlExam = @"
    SELECT * FROM exam";

            DataTable dtExam = qh.Select(sqlExam);

            foreach (DataRow row in dtExam.Rows)
            {
                string exam = "" + row["exam_name"];

                _examList.Add(exam);
            }
            #endregion

            #region 抓領域

            bkw.ReportProgress(60);
            // 取得 教務作業的設定領域名稱
            string ConfigName = "JHEvaluation_Subject_Ordinal";
            string ColumnKey = "DomainOrdinal";

            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration[ConfigName];
            if (cd.Contains(ColumnKey))
            {
                XmlElement element = cd.GetXml(ColumnKey, XmlHelper.LoadXml("<Domains/>"));
                foreach (XmlElement domainElement in element.SelectNodes("Domain"))
                {
                    string group = domainElement.GetAttribute("Group");
                    string name = domainElement.GetAttribute("Name");
                    string englishName = domainElement.GetAttribute("EnglishName");

                    _doaminList.Add(name);
                }

            }
            _doaminList.Add(""); //有些課程可能會沒有領域， 給它空值。
            #endregion


            bkw.ReportProgress(80);

            FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

            // 因原本樣板儲存 是 跟隨ESL 設定樣板， 現在取消掉，全部都要獨立設定
            // 因此 沒有參考ESL 樣板ID 的都是 學生ESL 成績單 的樣板設定檔
            string qry = "Ref_exam_Template_ID is null";

            _configuresList = _AccessHelper.Select<UDT_ReportTemplate>(qry);



            bkw.ReportProgress(100);

        }

        void bkw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            circularProgress1.Value = e.ProgressPercentage;
        }

        void bkw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 加入 科目名稱
            foreach (string subject in _subjectList)
            {
                comboBoxEx3.Items.Add(subject);
            }

            // 進度條 隱藏
            circularProgress1.Hide();

            cboConfigure.Items.Clear();
            foreach (var item in _configuresList)
            {
                cboConfigure.Items.Add(item);
            }
            cboConfigure.Items.Add(new UDT_ReportTemplate() { Name = "新增" });

        }



        // 列印
        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (dtBegin.IsEmpty || dtEnd.IsEmpty)
            {
                FISCA.Presentation.Controls.MsgBox.Show("日期區間必須輸入!");
                return;
            }

            if (dtBegin.Value > dtEnd.Value)
            {
                FISCA.Presentation.Controls.MsgBox.Show("開始日期必須小於或等於結束日期!!");
                return;
            }

            _BeginDate = dtBegin.Value;
            _EndDate = dtEnd.Value;

            // 將時間轉為本地時間，以防語系時間設定問題
            _BeginDate.ToShortDateString();
            _EndDate.ToShortDateString();

            // 關閉畫面控制項
            lnkCopyConfig.Enabled = false;
            lnkDelConfig.Enabled = false;
            btnPrint.Enabled = false;
            btnClose.Enabled = false;
            linklabel1.Enabled = false;
            linklabel2.Enabled = false;
            linklabel3.Enabled = false;
            linkLabel4.Enabled = false;

            school_year = comboBoxEx1.Text;
            semester = comboBoxEx2.Text;

            // 假如使用者有選擇科目，列印順序將以 課程科目排序、而非班級學生
            if (comboBoxEx3.Text != "")
            {
                _isOrderBySubject = true;
                _orderedSubject = comboBoxEx3.Text; // 排序科目
            }

            FormParam formParam = new FormParam(school_year, semester);

            _bw.RunWorkerAsync(formParam);

            //2018/12/20 不選電腦內樣板了，有設定檔
            //// 2018/10/29 穎驊註解，目前的作法先每一次都讓使用者列印前，選擇列印樣板
            //// 等本次期中考後，再看使用情境 怎麼去做列印樣板設定。
            //OpenFileDialog ope = new OpenFileDialog();
            //ope.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";

            //if (ope.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            //{
            //    return;
            //}
            //else
            //{
            //    _wordURL = ope.FileName;
            //    _bw.RunWorkerAsync();

            //}
        }

        // 離開
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // 檢視套印樣板
        private void linklabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 當沒有設定檔
            if (_configure == null) return;
            linklabel1.Enabled = false;
            #region 儲存檔案

            string reportName = "ESL學生成績單樣板(" + _configure.Name + ").docx";

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".docx");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                System.IO.FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                _configure.Template.Save(stream, Aspose.Words.SaveFormat.Docx);

                stream.Flush();
                stream.Close();
                System.Diagnostics.Process.Start(path);
            }
            catch
            {

            }
            linklabel1.Enabled = true;
            #endregion
        }

        // 變更套印樣板
        private void linklabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_configure == null) return;
            linklabel2.Enabled = false;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Word檔案 (*.doc)|*.doc|Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    _configure.Template = new Aspose.Words.Document(dialog.FileName);
                    _configure.Encode();
                    _configure.Save();
                    MessageBox.Show("樣板上傳成功。");
                }
                catch
                {
                    MessageBox.Show("樣板開啟失敗。");
                }
            }
            linklabel2.Enabled = true;
        }

        // 檢視功能變數總表
        private void linklabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            // 新寫法，依照　所選取的ESL 樣板設定，產生出動態對應階層的功能變數總表！！
            CreateFieldTemplate();
            return;
        }


        // 穎驊搬過來的工具，可以一次大量建立有規則的功能變數，可以省下很多時間。
        private void CreateFieldTemplate()
        {
            // 儲放樣板 id 與 樣板名稱的對照
            Dictionary<string, string> templateIDNameDict = new Dictionary<string, string>();

            List<Term> termList = new List<Term>();

            Aspose.Words.Document doc = new Aspose.Words.Document();
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);

            // Create a paragraph style and specify some formatting for it.
            Aspose.Words.Style style = builder.Document.Styles.Add(Aspose.Words.StyleType.Paragraph, "ESLNameStyle");

            style.Font.Size = 24;
            style.Font.Bold = true;
            style.ParagraphFormat.SpaceAfter = 12;

            #region 固定變數

            builder.ParagraphFormat.Style = builder.Document.Styles["ESLNameStyle"];
            // 固定變數，不分　期中、期末、學期  (使用大字粗體)
            builder.Writeln("固定變數");

            builder.ParagraphFormat.Style = builder.Document.Styles["Normal"];

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
                    "原班級名稱",
                    "學生英文姓名",
                    "學生中文姓名",
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


            #region 日常生活表現統計

            builder.ParagraphFormat.Style = builder.Document.Styles["ESLNameStyle"];
            // 固定變數，不分　期中、期末、學期  (使用大字粗體)
            builder.Writeln("日常生活表現統計");

            builder.ParagraphFormat.Style = builder.Document.Styles["Normal"];

            builder.StartTable();
            builder.InsertCell();
            builder.Write("項目");
            builder.InsertCell();
            builder.Write("變數");
            builder.EndRow();

            List<string> itemList = new List<string>();

            itemList.Add("區間開始日期");
            itemList.Add("區間結束日期");
            itemList.Add("大功區間統計");
            itemList.Add("小功區間統計");
            itemList.Add("嘉獎區間統計");
            itemList.Add("大過區間統計");
            itemList.Add("小過區間統計");
            itemList.Add("警告區間統計");


            // 缺曠欄位
            foreach (var type in _typeList)
            {
                foreach (var absence in _absenceList)
                {
                    itemList.Add(type + absence);
                }
            }

            foreach (string key in itemList
                )
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


            #region  解讀　description　XML
            // 取得ESL 描述 in description
            DataTable dt;
            QueryHelper qh = new QueryHelper();

            // 抓所有目前系統 有設定的ESL 評分樣版
            string selQuery = "SELECT id,name,description FROM exam_template WHERE description IS NOT NULL ORDER BY name  ";
            dt = qh.Select(selQuery);

            foreach (DataRow dr in dt.Rows)
            {
                string xmlStr = "<root>" + dr["description"].ToString() + "</root>";

                string eslTemplateName = dr["name"].ToString();

                XElement elmRoot = XElement.Parse(xmlStr);

                termList = ElmRootToTermlist(elmRoot);

                MergeFieldGenerator(builder, eslTemplateName, termList);

                termList.Clear(); // 每一個樣板 清完後 再加

                templateIDNameDict.Add(dr["id"].ToString(), dr["name"].ToString());
            }

            // 幫每一個系統試別 建立序列化的功能變數 
            foreach (string exam in _examList)
            {
                MergeFieldSerialGenerator(builder, exam, _doaminList);
            }

            #region 課程總成績的序列化功能變數
            // 課程總成績的序列化功能變數
            // Apply the paragraph style to the current paragraph in the document and add some text.
            builder.ParagraphFormat.Style = builder.Document.Styles["ESLNameStyle"];
            // 每一個 ESL 樣板的名稱 放在最上面 (使用大字粗體)
            builder.Writeln("課程總成績:");

            // Change to a paragraph style that has no list formatting. (將字體還原)
            builder.ParagraphFormat.Style = builder.Document.Styles["Normal"];
            builder.Writeln("");

            builder.StartTable();

            builder.InsertCell();
            builder.Write("學期科目成績算術平均");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "學期科目成績算術平均" + " \\* MERGEFORMAT ", "«SS_AVG»");
            builder.EndRow();



            builder.InsertCell();
            builder.Write("學期科目GPA算術平均");

            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "學期科目GPA算術平均" + " \\* MERGEFORMAT ", "«SSGPA_AVG»");
            builder.EndRow();


            builder.InsertCell();
            builder.Write("課程名稱");
            builder.InsertCell();
            builder.Write("課程科目名稱");
            builder.InsertCell();
            builder.Write("課程學期成績分數");
            builder.InsertCell();
            builder.Write("課程學期成績等第");
            builder.InsertCell();
            builder.Write("課程學期成績GPA");
            builder.InsertCell();
            builder.Write("課程文字描述");

            builder.EndRow();

            for (int i = 1; i < 26; i++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "課程名稱" + i + " \\* MERGEFORMAT ", "«CN»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "課程科目名稱" + i + " \\* MERGEFORMAT ", "«CSN»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "課程學期成績分數" + i + " \\* MERGEFORMAT ", "«CSS»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "課程學期成績等第" + i + " \\* MERGEFORMAT ", "«CSL»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "課程學期成績GPA" + i + " \\* MERGEFORMAT ", "«CSL»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "課程文字描述" + i + " \\* MERGEFORMAT ", "«CTEXT»"); // 20200615客製
                builder.EndRow();
            }

            builder.EndTable();
            builder.Writeln();
            #endregion





            #endregion


            #region 儲存檔案
            string inputReportName = "合併欄位總表";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".docx");

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
                doc.Save(path, Aspose.Words.SaveFormat.Docx);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".docx";
                sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        doc.Save(path, Aspose.Words.SaveFormat.Docx);
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

        private List<Term> ElmRootToTermlist(XElement elmRoot)
        {
            List<Term> termList = new List<Term>();

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

                        termList.Add(t); // 整理成大包的termList 後面用此來拚功能變數總表
                    }
                }
            }

            return termList;
        }

        private void MergeFieldGenerator(Aspose.Words.DocumentBuilder builder, string eslTemplateName, List<Term> termList)
        {
            #region 成績變數

            int termCounter = 1;

            // 2018/6/15 穎驊備註 以下整理 功能變數 最常使用的 string..Trim().Replace(' ', '_').Replace('"', '_') 
            // >> 其用意為避免Word 功能變數合併列印時 會有一些奇怪的BUG ，EX: row["Final-Term評量_Science科目_In-Class Score子項目_分數1"] = "YOYO!"; >> 有空格印不出來 


            // Apply the paragraph style to the current paragraph in the document and add some text.
            builder.ParagraphFormat.Style = builder.Document.Styles["ESLNameStyle"];
            // 每一個 ESL 樣板的名稱 放在最上面 (使用大字粗體)
            builder.Writeln("樣板名稱: " + eslTemplateName);

            // Change to a paragraph style that has no list formatting. (將字體還原)
            builder.ParagraphFormat.Style = builder.Document.Styles["Normal"];
            builder.Writeln("");

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

                builder.Write(term.Name);

                builder.InsertCell();

                // 2018/10/29 穎驊註解，和恩正討論後，不同樣板之間的 Term 名稱 會分不清楚， 因此在前面加 評分樣板作區別
                builder.InsertField("MERGEFIELD " + eslTemplateName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數" + " \\* MERGEFORMAT ", "«TS»");

                builder.InsertCell();

                builder.InsertField("MERGEFIELD " + eslTemplateName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重" + " \\* MERGEFORMAT ", "«TW»");

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
                    builder.InsertCell();
                    builder.Write("教師");

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
                        builder.Write(assessment.Name);

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重" + " \\* MERGEFORMAT ", "«AW»");

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數" + " \\* MERGEFORMAT ", "«AS»");

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師" + " \\* MERGEFORMAT ", "«AT»");

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
                        builder.InsertCell();
                        builder.Write("教師");
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
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師" + " \\* MERGEFORMAT ", "«T»");
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
                        builder.InsertCell();
                        builder.Write("教師");
                        builder.EndRow();

                        assessmentCounter = 1;
                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            if (assessment.Type == "Comment") // 檢查看有沒有　Comment　，專為 Comment 畫張表
                            {
                                builder.InsertCell();
                                builder.Write(assessment.Name);
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "評語" + " \\* MERGEFORMAT ", "«C»");
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師" + " \\* MERGEFORMAT ", "«T»");
                                builder.EndRow();
                                assessmentCounter++;
                            }
                        }
                        builder.EndTable();
                        builder.Writeln();
                    }

                }


                builder.Writeln();
                builder.Writeln();
            }


            #region 課程學期成績
            builder.Writeln("課程學期成績");

            builder.StartTable();
            builder.InsertCell();
            builder.Write("課程學期成績分數");
            builder.InsertCell();
            builder.Write("課程學期成績等第");
            builder.EndRow();

            builder.InsertCell();

            builder.InsertField("MERGEFIELD " + eslTemplateName.Replace(' ', '_').Replace('"', '_') + "_" + "課程學期成績分數" + " \\* MERGEFORMAT ", "«CSS»");

            builder.InsertCell();

            builder.InsertField("MERGEFIELD " + eslTemplateName.Replace(' ', '_').Replace('"', '_') + "_" + "課程學期成績等第" + " \\* MERGEFORMAT ", "«CSL»");

            builder.EndRow();
            builder.EndTable();
            #endregion

            #endregion

        }

        // 產生序列化的 功能變數
        private void MergeFieldSerialGenerator(Aspose.Words.DocumentBuilder builder, string ExamName, List<string> doaminList)
        {
            #region 依領域分列


            // 2018/6/15 穎驊備註 以下整理 功能變數 最常使用的 string..Trim().Replace(' ', '_').Replace('"', '_') 
            // >> 其用意為避免Word 功能變數合併列印時 會有一些奇怪的BUG ，EX: row["Final-Term評量_Science科目_In-Class Score子項目_分數1"] = "YOYO!"; >> 有空格印不出來 


            // Apply the paragraph style to the current paragraph in the document and add some text.
            builder.ParagraphFormat.Style = builder.Document.Styles["ESLNameStyle"];
            // 每一個 評量的名稱 放在最上面 (使用大字粗體)
            builder.Writeln("評量名稱: " + ExamName + " 依領域分列");

            // Change to a paragraph style that has no list formatting. (將字體還原)
            builder.ParagraphFormat.Style = builder.Document.Styles["Normal"];
            builder.Writeln("");
            foreach (string domain in doaminList)
            {
                // 領域
                builder.Writeln(ExamName + "_" + domain.Trim().Replace(' ', '_').Replace('"', '_'));

                builder.StartTable();
                builder.InsertCell();
                builder.Write("課程科目名稱");
                builder.InsertCell();
                builder.Write("課程教師一");
                builder.InsertCell();
                builder.Write("科目權數");
                builder.InsertCell();
                builder.Write("科目定期評量");
                builder.InsertCell();
                builder.Write("科目平時評量");
                builder.InsertCell();
                builder.Write("科目總成績");
                builder.InsertCell();
                builder.Write("科目文字描述");
                //builder.InsertCell();
                //builder.Write("科目文字評量");
                //builder.InsertCell();
                //builder.Write("文字描述");
                builder.EndRow();

                for (int i = 1; i < 8; i++)
                {
                    builder.InsertCell();

                    builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "課程科目名稱" + i + " \\* MERGEFORMAT ", "«SN»");

                    builder.InsertCell();

                    builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "課程教師一" + i + " \\* MERGEFORMAT ", "«ST»");

                    builder.InsertCell();

                    builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "科目權數" + i + " \\* MERGEFORMAT ", "«SC»");

                    builder.InsertCell();

                    builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "科目定期評量" + i + " \\* MERGEFORMAT ", "«SF»");

                    builder.InsertCell();

                    builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "科目平時評量" + i + " \\* MERGEFORMAT ", "«SA»");

                    builder.InsertCell();

                    builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "科目總成績" + i + " \\* MERGEFORMAT ", "«SST»");


                    builder.InsertCell();

                    builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "科目文字描述" + i + " \\* MERGEFORMAT ", "«STEXT»");

                    builder.EndRow();
                }

                builder.EndTable();
                builder.Writeln();
            }



            #endregion

            #region 所有科目


            // 2018/6/15 穎驊備註 以下整理 功能變數 最常使用的 string..Trim().Replace(' ', '_').Replace('"', '_') 
            // >> 其用意為避免Word 功能變數合併列印時 會有一些奇怪的BUG ，EX: row["Final-Term評量_Science科目_In-Class Score子項目_分數1"] = "YOYO!"; >> 有空格印不出來 


            // Apply the paragraph style to the current paragraph in the document and add some text.
            builder.ParagraphFormat.Style = builder.Document.Styles["ESLNameStyle"];
            // 每一個 評量的名稱 放在最上面 (使用大字粗體)
            builder.Writeln("評量名稱: " + ExamName + " 所有科目");

            // Change to a paragraph style that has no list formatting. (將字體還原)
            builder.ParagraphFormat.Style = builder.Document.Styles["Normal"];
            builder.Writeln("");

            builder.StartTable();
            builder.InsertCell();
            builder.Write("課程科目名稱");
            builder.InsertCell();
            builder.Write("課程教師一");
            builder.InsertCell();
            builder.Write("科目權數");
            builder.InsertCell();
            builder.Write("科目定期評量");
            builder.InsertCell();
            builder.Write("科目平時評量");
            builder.InsertCell();
            builder.Write("科目總成績");
            builder.InsertCell();
            builder.Write("科目文字描述");
            builder.EndRow();

            for (int i = 1; i < 26; i++)
            {
                builder.InsertCell();

                builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_課程科目名稱" + i + " \\* MERGEFORMAT ", "«SN»");

                builder.InsertCell();

                builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_課程教師一" + i + " \\* MERGEFORMAT ", "«ST»");

                builder.InsertCell();

                builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_科目權數" + i + " \\* MERGEFORMAT ", "«SC»");

                builder.InsertCell();

                builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_科目定期評量" + i + " \\* MERGEFORMAT ", "«SF»");

                builder.InsertCell();

                builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_科目平時評量" + i + " \\* MERGEFORMAT ", "«SA»");

                builder.InsertCell();

                builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_科目總成績" + i + " \\* MERGEFORMAT ", "«SST»");


                builder.InsertCell();

                builder.InsertField("MERGEFIELD " + ExamName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_科目文字描述" + i + " \\* MERGEFORMAT ", "«SST»");




                builder.EndRow();
            }

            builder.EndTable();
            builder.Writeln();



            #endregion
        }

        private void _bkWork_DoWork(object sender, DoWorkEventArgs e)
        {
            FormParam formParam = e.Argument as FormParam;

            // 處理等第
            DegreeMapper dm = new DegreeMapper();

            List<string> studentIDList = new List<string>();
            List<string> courseIDList = new List<string>();

            // 選擇的學生名單 
            studentIDList = K12.Presentation.NLDPanels.Student.SelectedSource;


            #region 取得本學期有設定ESL 評分樣版的課程清單
            QueryHelper qh = new QueryHelper();

            // 選出 本學期有設定ESL 評分樣版的清單
            string sqlCourseID = @"
    SELECT 
		id
		,course_name
		,school_year
		,semester
		,ref_exam_template_id 
	FROM course 
	WHERE 
	school_year =" + school_year +
    @"AND semester = " + semester +
    @"AND ref_exam_template_id IN(
		SELECT id 
		FROM exam_template  
		WHERE description IS NOT NULL)";

            DataTable dtCourseID = qh.Select(sqlCourseID);

            foreach (DataRow row in dtCourseID.Rows)
            {
                string id = "" + row["id"];

                courseIDList.Add(id);
            }

            if (courseIDList.Count == 0)
            {
                e.Cancel = true;
                MsgBox.Show("本學期沒有任何設定ESL樣板的課程。", "錯誤!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion


            #region  解讀　description　XML

            // 取得ESL 描述 in description
            DataTable dtTemplateDescription;

            string selQuery = "SELECT id,name,description FROM exam_template WHERE description IS NOT NULL";

            dtTemplateDescription = qh.Select(selQuery);

            Dictionary<string, List<Term>> termListDict = new Dictionary<string, List<Term>>();

            List<List<Term>> termListCollection = new List<List<Term>>();

            //整理目前的ESL 課程資料
            if (dtTemplateDescription.Rows.Count > 0)
            {
                foreach (DataRow dr in dtTemplateDescription.Rows)
                {
                    List<Term> termList = new List<Term>();

                    string xmlStr = "<root>" + dr["description"].ToString() + "</root>";
                    XElement elmRoot = XElement.Parse(xmlStr);

                    termList = ElmRootToTermlist(elmRoot);

                    termListDict.Add(dr["name"].ToString(), termList);

                }

                _mergeDataTable = GetMergeField(termListDict);
            }


            #endregion

            // 取得課程基本資料 (教師)
            List<K12.Data.CourseRecord> courseList = K12.Data.Course.SelectByIDs(courseIDList);

            // 取得教師基本資料 
            List<K12.Data.TeacherRecord> teacherList = K12.Data.Teacher.SelectAll();

            //取得學生基本資料
            List<K12.Data.StudentRecord> studentList = K12.Data.Student.SelectByIDs(studentIDList);


            #region 取得ESL 課程成績
            _bw.ReportProgress(20, "取得ESL課程成績");

            string course_ids = string.Join("','", courseIDList);

            string student_ids = string.Join("','", studentIDList);

            // 建立成績結構
            foreach (string stuID in studentIDList)
            {
                _scoreDict.Add(stuID, new Dictionary<string, string>());
            }

            // (舊的SQL 按照ESL 2019寒假優化計畫 ， 全面採用 ref_sc_attend_id 欄位)
            // 按照時間順序抓， 如果有相同的成績結構， 以後來新的 取代前的
            //string sqlScore = @"    
            //SELECT 
            //        $esl.gradebook_assessment_score.last_update
            //        ,$esl.gradebook_assessment_score.term
            //        ,$esl.gradebook_assessment_score.subject
            //        ,$esl.gradebook_assessment_score.assessment
            //        ,$esl.gradebook_assessment_score.custom_assessment
            //        ,$esl.gradebook_assessment_score.ref_course_id
            //        ,$esl.gradebook_assessment_score.ref_student_id
            //        ,$esl.gradebook_assessment_score.ref_teacher_id
            //        ,$esl.gradebook_assessment_score.ref_sc_attend_id
            //        ,$esl.gradebook_assessment_score.value
            //        ,$esl.gradebook_assessment_score.ratio  
            //        ,exam_template.name AS exam_template_name
            //FROM $esl.gradebook_assessment_score    
            //        LEFT JOIN course ON course.id =$esl.gradebook_assessment_score.ref_course_id
            //        LEFT JOIN exam_template ON exam_template.id =  course.ref_exam_template_id
            //WHERE 
            //        ref_course_id IN ('" + course_ids + @"') 
            //        AND ref_student_id IN('" + student_ids + @"')
            //ORDER BY last_update,ref_student_id ";

            // 新SQL 採用  ref_sc_attend_id 欄位
            // 按照時間順序抓， 如果有相同的成績結構， 以後來新的 取代前的
            string sqlScore = @"    
            SELECT 
                    $esl.gradebook_assessment_score.last_update
                    ,$esl.gradebook_assessment_score.term
                    ,$esl.gradebook_assessment_score.subject
                    ,$esl.gradebook_assessment_score.assessment
                    ,$esl.gradebook_assessment_score.custom_assessment
                    ,sc_attend.ref_course_id
                    ,sc_attend.ref_student_id
                    ,$esl.gradebook_assessment_score.ref_teacher_id
                    ,$esl.gradebook_assessment_score.ref_sc_attend_id
                    ,$esl.gradebook_assessment_score.value
                    ,$esl.gradebook_assessment_score.ratio  
                    ,exam_template.name AS exam_template_name
            FROM $esl.gradebook_assessment_score    
                    LEFT JOIN sc_attend ON sc_attend.id =$esl.gradebook_assessment_score.ref_sc_attend_id
                    LEFT JOIN course ON course.id =sc_attend.ref_course_id
                    LEFT JOIN exam_template ON exam_template.id =  course.ref_exam_template_id
            WHERE 
                    sc_attend.ref_course_id IN ('" + course_ids + @"') 
                    AND sc_attend.ref_student_id IN('" + student_ids + @"')
            ORDER BY $esl.gradebook_assessment_score.last_update,sc_attend.ref_student_id ";

            DataTable dtScore = qh.Select(sqlScore);

            decimal progress = 20;
            decimal per = (decimal)(100 - progress) / (dtScore.Rows.Count != 0 ? dtScore.Rows.Count : 1);

            foreach (DataRow row in dtScore.Rows)
            {
                progress += (decimal)(per);
                _bw.ReportProgress((int)progress);

                string termWord = "" + row["term"];
                string subjectWord = "" + row["subject"];
                string assessmentWord = "" + row["assessment"];

                string examTemplateName = "" + row["exam_template_name"];

                string id = "" + row["ref_student_id"];

                // 有教師自訂的子項目成績就跳掉 不處理
                if ("" + row["custom_assessment"] != "")
                {
                    continue;
                }

                // 項目都有，為assessment 成績
                if (termWord != "" && "" + subjectWord != "" && "" + assessmentWord != "")
                {
                    if (_scoreDict.ContainsKey(id))
                    {
                        // 指標型成績
                        if (_indicatorList.Contains("" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_')))
                        {
                            string scoreKey = "評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_');

                            if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "指標"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "指標", "" + row["value"]);
                            }
                            else
                            {
                                _scoreDict[id][scoreKey + "_" + "指標"] = "" + row["value"];
                            }

                            if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "教師"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "教師", teacherList.Find(t => t.ID == "" + row["ref_teacher_id"]).Name); //教師名稱
                            }
                            else
                            {
                                _scoreDict[id][scoreKey + "_" + "教師"] = teacherList.Find(t => t.ID == "" + row["ref_teacher_id"]).Name; //教師名稱
                            }

                        }
                        // 評語型成績
                        else if (_commentList.Contains("" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_')))
                        {
                            string scoreKey = "評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_');

                            if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "評語"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "評語", "" + row["value"]);
                            }
                            else
                            {
                                _scoreDict[id][scoreKey + "_" + "評語"] = "" + row["value"];
                            }

                            if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "教師"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "教師", teacherList.Find(t => t.ID == "" + row["ref_teacher_id"]).Name); //教師名稱
                            }
                            else
                            {
                                _scoreDict[id][scoreKey + "_" + "教師"] = teacherList.Find(t => t.ID == "" + row["ref_teacher_id"]).Name; //教師名稱
                            }

                        }
                        // 分數型成績
                        else
                        {
                            string scoreKey = "評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_');

                            if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "分數"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "分數", "" + row["value"]);

                                if (_itemDict[examTemplateName].ContainsKey(scoreKey + "_" + "比重"))
                                {
                                    _scoreDict[id].Add(scoreKey + "_" + "比重", _itemDict[examTemplateName][scoreKey + "_" + "比重"]);
                                }
                            }
                            else
                            {
                                _scoreDict[id][scoreKey + "_" + "分數"] = "" + row["value"];
                            }

                            if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "教師"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "教師", teacherList.Find(t => t.ID == "" + row["ref_teacher_id"]).Name); //教師名稱
                            }
                            else
                            {
                                _scoreDict[id][scoreKey + "_" + "教師"] = teacherList.Find(t => t.ID == "" + row["ref_teacher_id"]).Name; //教師名稱
                            }

                        }
                    }
                }

                // 沒有assessment，為subject 成績
                if (termWord != "" && "" + subjectWord != "" && "" + assessmentWord == "")
                {
                    if (_scoreDict.ContainsKey(id))
                    {
                        string scoreKey = "評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_');

                        if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "分數"))
                        {
                            _scoreDict[id].Add(scoreKey + "_" + "分數", "" + row["value"]);

                            if (_itemDict[examTemplateName].ContainsKey(scoreKey + "_" + "比重"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "比重", _itemDict[examTemplateName][scoreKey + "_" + "比重"]);
                            }
                        }
                        else
                        {
                            _scoreDict[id][scoreKey + "_" + "分數"] = "" + row["value"];
                        }

                    }
                }
                // 沒有assessment、subject，為term 成績
                if (termWord != "" && "" + subjectWord == "" && "" + assessmentWord == "")
                {

                    if (_scoreDict.ContainsKey(id))
                    {
                        // 2018/10/29 穎驊註解，和恩正討論後，不同樣板之間的 Term 名稱 會分不清楚， 因此在前面加 評分樣板作區別
                        string scoreKey = examTemplateName.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_');

                        if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "分數"))
                        {
                            _scoreDict[id].Add(scoreKey + "_" + "分數", "" + row["value"]);

                            if (_itemDict[examTemplateName].ContainsKey(scoreKey + "_" + "比重"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "比重", _itemDict[examTemplateName][scoreKey + "_" + "比重"]);
                            }
                        }
                        else
                        {
                            _scoreDict[id][scoreKey + "_" + "分數"] = "" + row["value"];
                        }
                    }
                }

            }


            //2018/12/18 穎驊與恩正討論後， 對應 序列化 成績的功能變數， 從學生 系統評量成績 sce_take 抓取  對應成績
            string sqlSerialScore = @"
SELECT 
exam.exam_name
--,exam_template.extension AS ratio
,unnest(xpath('/Extension/ScorePercentage/text()', xmlparse(content exam_template.extension))) ::text AS exam_ratio
,100 - (unnest(xpath('/Extension/ScorePercentage/text()', xmlparse(content exam_template.extension))) ::text)::integer AS assignment_ratio
,teacher.teacher_name
,course.credit
,course.domain
,course.subject
,sc_attend.ref_student_id
,sce_take.score AS score -- 評量總分
--,sce_take.extension
,array_to_string(xpath('/Extension/Score/text()', xmlparse(content sce_take.extension)),'') ::text AS exam_score 
,array_to_string(xpath('/Extension/AssignmentScore/text()', xmlparse(content sce_take.extension)),'') ::text AS assignment_score 
,array_to_string(xpath('/Extension/Text/text()', xmlparse(content sce_take.extension)),'') ::text AS text 
FROM sce_take 
LEFT JOIN sc_attend ON sc_attend.id = sce_take.ref_sc_attend_id
LEFT JOIN course ON course.id = sc_attend.ref_course_id
LEFT JOIN exam ON exam.id =sce_take.ref_exam_id
LEFT JOIN exam_template ON exam_template.id =course.ref_exam_template_id
LEFT JOIN tc_instruct ON course.id =tc_instruct.ref_course_id
LEFT JOIN teacher ON teacher.id =tc_instruct.ref_teacher_id
WHERE course.id IN ('" + course_ids + "') " +
"AND sc_attend.ref_student_id IN ('" + student_ids + "')" +
"AND tc_instruct.sequence =1" +
"ORDER BY ref_student_id,exam_name,domain,subject";


            DataTable dtSerialScore = qh.Select(sqlSerialScore);



            foreach (DataRow row in dtSerialScore.Rows)
            {
                string id = "" + row["ref_student_id"];

                string examWord = "" + row["exam_name"];
                string domainWord = "" + row["domain"];
                string subjectWord = "" + row["subject"];
                string teacher = "" + row["teacher_name"]; // 教師姓名，固定抓 該課程的 教師一
                string credit = "" + row["credit"];
                string exam_score = "" + row["exam_score"]; // 定期評量
                string assignment_score = "" + row["assignment_score"]; // 平時評量
                string score = "" + row["score"]; // 評量總分
                string text = "" + row["text"];
                {//依領域分開
                    string scoreKey = examWord.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domainWord + "_";

                    if (_scoreDict.ContainsKey(id))
                    {
                        bool added = false;  // 尚未加入

                        for (int i = 1; !added; i++)
                        {
                            // 課程科目名稱
                            if (_scoreDict[id].ContainsKey(scoreKey + "課程科目名稱" + i))
                            {
                                continue;
                            }
                            else
                            {
                                _scoreDict[id].Add(scoreKey + "課程科目名稱" + i, subjectWord);
                                added = true;
                            }

                            _scoreDict[id].Add(scoreKey + "課程教師一" + i, teacher);

                            _scoreDict[id].Add(scoreKey + "科目權數" + i, credit);

                            _scoreDict[id].Add(scoreKey + "科目定期評量" + i, exam_score);

                            _scoreDict[id].Add(scoreKey + "科目平時評量" + i, assignment_score);

                            _scoreDict[id].Add(scoreKey + "科目總成績" + i, score);

                            _scoreDict[id].Add(scoreKey + "科目文字描述" + i, text);

                        }
                    }
                }
                {//所有科目
                    string scoreKey = examWord.Replace(' ', '_').Replace('"', '_') + "_" + "評量_";

                    if (_scoreDict.ContainsKey(id))
                    {
                        bool added = false;  // 尚未加入

                        for (int i = 1; !added; i++)
                        {
                            // 課程科目名稱
                            if (_scoreDict[id].ContainsKey(scoreKey + "課程科目名稱" + i))
                            {
                                continue;
                            }
                            else
                            {
                                _scoreDict[id].Add(scoreKey + "課程科目名稱" + i, subjectWord);
                                added = true;
                            }

                            _scoreDict[id].Add(scoreKey + "課程教師一" + i, teacher);

                            _scoreDict[id].Add(scoreKey + "科目權數" + i, credit);

                            _scoreDict[id].Add(scoreKey + "科目定期評量" + i, exam_score);

                            _scoreDict[id].Add(scoreKey + "科目平時評量" + i, assignment_score);

                            _scoreDict[id].Add(scoreKey + "科目總成績" + i, score);

                            _scoreDict[id].Add(scoreKey + "科目文字描述" + i, text);

                        }
                    }
                }
            }





            // 課程學期成績
            string sqlSemesterCourseScore = @"SELECT
sc_attend.ref_student_id
,exam_template.name
,course.course_name
,course.domain
,course.semester
,course.school_year
,course.subject
,sc_attend.score
FROM sc_attend
LEFT JOIN course ON course.id = sc_attend.ref_course_id
LEFT JOIN exam_template ON exam_template.id =course.ref_exam_template_id
WHERE course.id IN ('" + course_ids + "') " +
"AND sc_attend.ref_student_id IN ('" + student_ids + "')" +
"ORDER BY ref_student_id,domain,subject";

            DataTable dtSemesterCourseScore = qh.Select(sqlSemesterCourseScore);

            DataService dataService = new DataService();

            // 整理科目文字評量 (【sems_subj_score 】) 
            Dictionary<string, Dictionary<string, string>> dicSubjectTexts = dataService.GetSemsSubjText(studentIDList, formParam.SchoolYear, formParam.Semester);           


            foreach (DataRow row in dtSemesterCourseScore.Rows)
            {
                string id = "" + row["ref_student_id"];

                string templateWord = "" + row["name"];
                string domainWord = "" + row["domain"];
                string courseWord = "" + row["course_name"];
                string subjectWord = "" + row["subject"];
                string score = "" + row["score"]; // 課程學期成績

                string scoreKey = templateWord.Trim().Replace(' ', '_').Replace('"', '_');




                if (_scoreDict.ContainsKey(id))
                {
                    #region 跟樣板的功能變數
                    // 理論上一學期上 一個學生 只會有一個ESL評分樣版的課程成績 ， 不會有同一個ESL 評分樣版 有不同的課程成績
                    if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "課程學期成績分數"))
                    {
                        _scoreDict[id].Add(scoreKey + "_" + "課程學期成績分數", score);
                    }
                    if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "課程學期成績等第"))
                    {
                        decimal score_d;
                        if (decimal.TryParse(score, out score_d))
                        {
                            _scoreDict[id].Add(scoreKey + "_" + "課程學期成績等第", dm.GetDegreeByScore(score_d));
                        }
                    }

                    #endregion
                    #region 序列化功能變數
                    bool added = false;  // 尚未加入
                
                    for (int i = 1; !added; i++)
                    {

                        // 課程名稱
                        if (_scoreDict[id].ContainsKey("課程名稱" + i))
                        {
                            continue;
                        }
                        else
                        {
                            _scoreDict[id].Add("課程名稱" + i, courseWord);
                        }

                        // 課程科目名稱
                        if (_scoreDict[id].ContainsKey("課程科目名稱" + i))
                        {
                            continue;
                        }
                        else
                        {
                            _scoreDict[id].Add("課程科目名稱" + i, subjectWord);
                            added = true;
                        }

                        _scoreDict[id].Add("課程學期成績分數" + i, score);




                        decimal score_d;

                        if (decimal.TryParse(score, out score_d))
                        {
                            _scoreDict[id].Add("課程學期成績等第" + i, dm.GetDegreeByScore(score_d));
                        }

                        decimal? GPA_d = dataService.GetFinalGPA(subjectWord, score_d);
                        _scoreDict[id].Add("課程學期成績GPA" + i, GPA_d.ToString());

                        // 處理學期算數平均
                        this._SemsTotalScoreInfos.AddforCoumpte(id, new SemsSubjScoreInfo(subjectWord, score_d, GPA_d));


                        if (_scoreDict[id].ContainsKey("課程文字描述" + i))
                        {
                            continue;
                        }
                        else
                        {
                            //不同來源之學期科目文字評量
                            if (dicSubjectTexts.ContainsKey(id))
                            {
                                if (dicSubjectTexts[id].ContainsKey(subjectWord))
                                {
                                    _scoreDict[id].Add("課程文字描述" + i, dicSubjectTexts[id][subjectWord]);
                                }
                            }
                        }


                    }
               
                        //decimal SemestAverageScore = Math.Round(totalScoreSum / subjCount, 1, MidpointRounding.AwayFromZero);
                        //_scoreDict[id].Add("學期科目成績算術平均", SemestAverageScore.ToString());
                        //decimal SemestAverageGPA = Math.Round(totalGPASum / subjCount, 1, MidpointRounding.AwayFromZero);
                        //_scoreDict[id].Add("學期科目GPA算術平均", SemestAverageGPA.ToString());
                 
                    #endregion
                }
            }

            #endregion

            //學期科目算數平均
            // 取得計算 學期算術平均的物件
            Dictionary<string, SemsTotalScoreInfo> semsScoreInfos =  this._SemsTotalScoreInfos.GetAllSemsTotalScoreInfo();

            foreach (string studID in this._SemsTotalScoreInfos.GetAllSemsTotalScoreInfo().Keys)
            {
                if (_scoreDict.ContainsKey(studID)) 
                {
                    SemsTotalScoreInfo semsTotalScoreInfo = semsScoreInfos[studID];
                    if (!_scoreDict[studID].ContainsKey("學期科目成績算術平均")) 
                    {
                    _scoreDict[studID].Add("學期科目成績算術平均", semsTotalScoreInfo.GetScoreAvg(2).ToString());
                    
                    }
                    if (!_scoreDict[studID].ContainsKey("學期科目GPA算術平均"))
                    {
                     _scoreDict[studID].Add("學期科目GPA算術平均", semsTotalScoreInfo.GetGPAAvg(2).ToString());
                    }
                }
            }





            #region 取得 缺曠獎懲

            _bw.ReportProgress(40, "取得缺曠資料");
            // 缺曠資料區間統計        
            _AttendanceDict = Utility.GetAttendanceCountByDate(studentList, _BeginDate, _EndDate);

            _bw.ReportProgress(50, "取得獎懲資料");
            // 獎懲資料
            _DisciplineCountDict = Utility.GetDisciplineCountByDate(studentIDList, _BeginDate, _EndDate);

            #endregion

            // BY subject 排序 ，將studentList 順序重新整理
            if (_isOrderBySubject)
            {
                // 2018/12/18 穎驊 備註 這一個API 有問題， 按下列所示，結果會抓不到東西。
                //List<K12.Data.SCAttendRecord> scaList = K12.Data.SCAttend.SelectByStudentIDAndCourseID(courseIDList, studentIDList);

                List<K12.Data.SCAttendRecord> scaList = K12.Data.SCAttend.SelectByCourseIDs(courseIDList);

                List<K12.Data.StudentRecord> studentList_new = new List<StudentRecord>();

                // 以課程ID 整理學生清單順序
                Dictionary<string, List<K12.Data.StudentRecord>> courseIDstudentListDict = new Dictionary<string, List<StudentRecord>>();


                foreach (SCAttendRecord scaRecord in scaList)
                {
                    if (scaRecord.Course.Subject == _orderedSubject)
                    {
                        StudentRecord stuRecord = studentList.Find(x => x.ID == scaRecord.RefStudentID);

                        if (stuRecord != null)
                        {
                            if (!courseIDstudentListDict.ContainsKey(scaRecord.Course.ID))
                            {
                                courseIDstudentListDict.Add(scaRecord.Course.ID, new List<StudentRecord>());

                                courseIDstudentListDict[scaRecord.Course.ID].Add(stuRecord);
                            }
                            else
                            {
                                courseIDstudentListDict[scaRecord.Course.ID].Add(stuRecord);
                            }

                            // 學生已經加入以科目的課程 新排序後， 從原本的清單移除
                            studentList.RemoveAll(s => s.ID == scaRecord.RefStudentID);
                        }
                    }
                }

                // 將每一個課程 內學生順序以學號排序
                foreach (string courseID in courseIDstudentListDict.Keys)
                {
                    courseIDstudentListDict[courseID].Sort((x, y) => { return x.StudentNumber.CompareTo(y.StudentNumber); });

                    // 排序後加入 新學生清單
                    studentList_new.AddRange(courseIDstudentListDict[courseID]);
                }

                // 剩下的學生 也用學號排序
                studentList.Sort((x, y) => { return x.StudentNumber.CompareTo(y.StudentNumber); });

                studentList_new.AddRange(studentList);

                studentList = studentList_new;
            }


            foreach (StudentRecord stuRecord in studentList)
            {
                string id = stuRecord.ID;

                DataRow row = _mergeDataTable.NewRow();

                row["電子報表辨識編號"] = "系統編號{" + stuRecord.ID + "}"; // 學生系統編號

                row["學年度"] = this.school_year;
                row["學期"] = this.semester;
                row["學號"] = stuRecord.StudentNumber;
                row["年級"] = stuRecord.Class != null ? "" + stuRecord.Class.GradeYear : "";

                row["原班級名稱"] = stuRecord.Class != null ? "" + stuRecord.Class.Name : "";
                row["學生英文姓名"] = stuRecord.EnglishName;
                row["學生中文姓名"] = stuRecord.Name;

                row["區間開始日期"] = _BeginDate.ToShortDateString();
                row["區間結束日期"] = _EndDate.ToShortDateString();


                // 傳入 ID當 Key
                // row["缺曠紀錄"] = StudRec.ID;                
                //缺曠套印
                foreach (var type in _typeList)
                {
                    foreach (var absence in _absenceList)
                    {
                        row[type + absence] = "0";
                    }
                }
                if (_AttendanceDict.ContainsKey(stuRecord.ID))
                {
                    foreach (var absentKey in _AttendanceDict[stuRecord.ID].Keys)
                    {
                        row[absentKey] = _AttendanceDict[stuRecord.ID][absentKey];
                    }
                }
                // 獎懲區間統計值
                if (_DisciplineCountDict.ContainsKey(stuRecord.ID))
                {
                    foreach (string str in Global.GetDisciplineNameList())
                    {
                        string key = str + "區間統計";
                        if (_DisciplineCountDict[stuRecord.ID].ContainsKey(str))
                            row[key] = _DisciplineCountDict[stuRecord.ID][str];
                    }
                }

                //foreach (string mergeKey in _itemDict.Keys)
                //{
                //    if (row.Table.Columns.Contains(mergeKey))
                //    {                                               
                //        row[mergeKey] = _itemDict[mergeKey];
                //    }
                //}

                if (_scoreDict.ContainsKey(id))
                {
                    foreach (string mergeKey in _scoreDict[id].Keys)
                    {
                        if (row.Table.Columns.Contains(mergeKey))
                        {
                            row[mergeKey] = _scoreDict[id][mergeKey];
                        }
                    }
                }

                _mergeDataTable.Rows.Add(row);

            }


            try
            {
                //// 載入使用者所選擇的 word 檔案
                //_doc = new Document(_wordURL);

                // 樣板的設定
                _doc = _configure.Template;
            }
            catch (Exception ex)
            {
                MsgBox.Show(ex.Message);
                e.Cancel = true;
                return;
            }


            _doc.MailMerge.Execute(_mergeDataTable);

            e.Result = _doc;
        }

        private void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                this.btnPrint.Enabled = true;
                this.btnPrint.Enabled = false;
                return;
            }

            FISCA.Presentation.MotherForm.SetStatusBarMessage(" ESL學生成績單產生完成。");

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
            sd.FileName = "ESL學生成績單.docx";
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

        private void _worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("" + e.UserState, e.ProgressPercentage);
        }

        private DataTable GetMergeField(Dictionary<string, List<Term>> termListDict)
        {
            // 計算權重用的字典(因為使用者在介面設定的權重數值 不一定就是想在報表上顯示的)
            // 目前康橋報表希望能夠將，每一個Subject、assessment 的比重 換算成為對於期中考的比例
            Dictionary<string, float> weightCalDict = new Dictionary<string, float>();

            DataTable dataTable = new DataTable();

            #region 固定變數
            // 固定變數
            // 基本資料
            dataTable.Columns.Add("學年度");
            dataTable.Columns.Add("學期");
            dataTable.Columns.Add("學號");
            dataTable.Columns.Add("年級");
            dataTable.Columns.Add("原班級名稱");
            dataTable.Columns.Add("學生英文姓名");
            dataTable.Columns.Add("學生中文姓名");
            dataTable.Columns.Add("電子報表辨識編號");

            dataTable.Columns.Add("區間開始日期");
            dataTable.Columns.Add("區間結束日期");


            #endregion


            // 獎懲名稱
            foreach (string str in Global.GetDisciplineNameList())
            {
                dataTable.Columns.Add(str + "區間統計");
            }

            // 缺曠欄位
            foreach (var type in _typeList)
            {
                foreach (var absence in _absenceList)
                {
                    dataTable.Columns.Add(type + absence);
                }
            }



            // 2018/6/15 穎驊備註 以下整理 功能變數 最常使用的 string.Trim().Replace(' ', '_').Replace('"', '_') 
            // >> 其用意為避免Word 功能變數合併列印時 會有一些奇怪的BUG ，EX: row["Final-Term評量_Science科目_In-Class Score子項目_分數1"] = "YOYO!"; >> 有空格印不出來 

            foreach (string templateName in termListDict.Keys)
            {
                //每一個 template 清空一次 weight 計算用 字典
                weightCalDict.Clear();

                _itemDict.Add(templateName, new Dictionary<string, string>());

                // 學期課程成績
                if (!dataTable.Columns.Contains(templateName.Replace(' ', '_').Replace('"', '_') + "_" + "課程學期成績分數"))
                    dataTable.Columns.Add(templateName.Replace(' ', '_').Replace('"', '_') + "_" + "課程學期成績分數");

                if (!dataTable.Columns.Contains(templateName.Replace(' ', '_').Replace('"', '_') + "_" + "課程學期成績等第"))
                    dataTable.Columns.Add(templateName.Replace(' ', '_').Replace('"', '_') + "_" + "課程學期成績等第");

                foreach (Term term in termListDict[templateName])
                {
                    // 2018/10/29 穎驊註解，和恩正討論後，不同樣板之間的 Term 名稱 會分不清楚， 因此在前面加 評分樣板作區別

                    if (!dataTable.Columns.Contains(templateName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重"))
                        dataTable.Columns.Add(templateName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重");

                    if (!dataTable.Columns.Contains(templateName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"))
                        dataTable.Columns.Add(templateName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"); // Term 分數本身 先暫時這樣處理之後要有類別整理

                    if (!_itemDict[templateName].ContainsKey(templateName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重"))
                    {
                        _itemDict[templateName].Add(templateName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重", term.Weight);
                    }


                    // 計算比重用，先整理 Term 、 Subject  的 總和
                    foreach (Subject subject in term.SubjectList)
                    {
                        // Term
                        if (!weightCalDict.ContainsKey(term.Name + "_SubjectTotal"))
                        {
                            if (float.TryParse(subject.Weight, out float f))
                            {
                                weightCalDict.Add(term.Name + "_SubjectTotal", f);
                            }
                        }
                        else
                        {
                            if (float.TryParse(subject.Weight, out float f))
                            {
                                weightCalDict[term.Name + "_SubjectTotal"] += f;
                            }
                        }

                        // Subject
                        if (!weightCalDict.ContainsKey(term.Name + "_" + subject.Name))
                        {
                            if (float.TryParse(subject.Weight, out float f))
                            {
                                weightCalDict.Add(term.Name + "_" + subject.Name, f);
                            }
                        }

                    }

                    foreach (Subject subject in term.SubjectList)
                    {
                        if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重"))
                            dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重");
                        if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"))
                            dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"); // Subject 分數本身 先暫時這樣處理之後要有類別整理


                        string subjectWieght = "" + Math.Round((float.Parse(subject.Weight) * 100) / (weightCalDict[term.Name + "_SubjectTotal"]), 2, MidpointRounding.ToEven);

                        var subjectKey = "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重";
                        if (!_itemDict[templateName].ContainsKey(subjectKey))
                            _itemDict[templateName].Add(subjectKey, subjectWieght); //subject比重 

                        // 計算比重用，先整理 Assessment  的 總和
                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            if (!weightCalDict.ContainsKey(term.Name + "_" + subject.Name + "_AssessmentTotal"))
                            {
                                if (float.TryParse(assessment.Weight, out float f))
                                {
                                    weightCalDict.Add(term.Name + "_" + subject.Name + "_AssessmentTotal", f);
                                }
                            }
                            else
                            {
                                if (float.TryParse(assessment.Weight, out float f))
                                {
                                    weightCalDict[term.Name + "_" + subject.Name + "_AssessmentTotal"] += f;
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
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師");

                                string assessmentWieght = "" + Math.Round((weightCalDict[term.Name + "_" + subject.Name] * float.Parse(assessment.Weight) * 100) / (weightCalDict[term.Name + "_SubjectTotal"] * weightCalDict[term.Name + "_" + subject.Name + "_AssessmentTotal"]), 2, MidpointRounding.ToEven);
                                var assessmentKey = "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重";
                                if (!_itemDict[templateName].ContainsKey(assessmentKey))
                                    _itemDict[templateName].Add(assessmentKey, assessmentWieght); //assessment比重 

                            }
                            if (assessment.Type == "Indicator") // 檢查看有沒有　　Indicator　，有的話另外存List 做對照
                            {
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "指標"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "指標");
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師");

                                string key = term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_');

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
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師");


                                string key = term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_');

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


            // 加入序列化的 功能變數
            foreach (string examName in _examList)
            {
                //依領域分開
                foreach (string domain in _doaminList)
                {
                    for (int i = 1; i < 8; i++)
                    {
                        dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "課程科目名稱" + i);
                        dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "課程教師一" + i);
                        dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "科目權數" + i);
                        dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "科目定期評量" + i);
                        dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "科目平時評量" + i);
                        dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "科目總成績" + i);
                        dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + domain + "_" + "科目文字描述" + i);
                    }
                }
                //所有科目
                for (int i = 1; i < 26; i++)
                {
                    dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_課程科目名稱" + i);
                    dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_課程教師一" + i);
                    dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_科目權數" + i);
                    dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_科目定期評量" + i);
                    dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_科目平時評量" + i);
                    dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_科目總成績" + i);
                    dataTable.Columns.Add(examName.Replace(' ', '_').Replace('"', '_') + "_" + "評量_科目文字描述" + i);
                }
            }
            dataTable.Columns.Add("學期科目成績算術平均");
            dataTable.Columns.Add("學期科目GPA算術平均");


            for (int i = 1; i < 26; i++)
            {
                dataTable.Columns.Add("課程名稱" + i);
                dataTable.Columns.Add("課程科目名稱" + i);
                dataTable.Columns.Add("課程學期成績分數" + i);
                dataTable.Columns.Add("課程學期成績GPA" + i);
                dataTable.Columns.Add("課程學期成績等第" + i);
                dataTable.Columns.Add("課程文字描述" + i);
            }


            return dataTable;
        }

        // ESL 成績 等第 設定
        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new ScoreMappingTable().ShowDialog();
        }

        private void PrintStudentESLReportForm_Load(object sender, EventArgs e)
        {
            bkw.RunWorkerAsync();
        }

        private void cboConfigure_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 選到最後一項 就是 新增樣板
            if (cboConfigure.SelectedIndex == cboConfigure.Items.Count - 1)
            {
                //新增
                btnPrint.Enabled = false;
                NewConfigure dialog = new NewConfigure();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    _configure = new UDT_ReportTemplate();
                    _configure.Name = dialog.ConfigName;
                    _configure.Template = dialog.Template;

                    _configuresList.Add(_configure);
                    cboConfigure.Items.Insert(cboConfigure.SelectedIndex, _configure);
                    cboConfigure.SelectedIndex = cboConfigure.SelectedIndex - 1;
                    _configure.Encode();
                    _configure.Save();
                    btnPrint.Enabled = true;
                }
                else
                {
                    cboConfigure.SelectedIndex = -1;
                }
            }
            else
            {
                if (cboConfigure.SelectedIndex >= 0)
                {
                    btnPrint.Enabled = true;
                    _configure = _configuresList[cboConfigure.SelectedIndex];
                    if (_configure.Template == null)
                        _configure.Decode();
                }
                else
                {
                    _configure = null;
                }

            }
        }

        // 刪除設定檔樣板
        private void lnkDelConfig_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_configure == null) return;

            if (MessageBox.Show("樣板刪除後將無法回復，確定刪除樣板?", "刪除樣板", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.OK)
            {
                _configuresList.Remove(_configure);
                if (_configure.UID != "")
                {
                    _configure.Deleted = true;
                    _configure.Save();
                }
                var conf = _configure;
                cboConfigure.SelectedIndex = -1;
                cboConfigure.Items.Remove(conf);
            }
        }

        // 複製評分樣版
        private void lnkCopyConfig_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_configure == null) return;
            CloneConfigure dialog = new CloneConfigure() { ParentName = _configure.Name };
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                UDT_ReportTemplate conf = new UDT_ReportTemplate();
                conf.Name = dialog.NewConfigureName;
                conf.Template = _configure.Template;
                conf.Encode();
                conf.Save();
                _configuresList.Add(conf);
                cboConfigure.Items.Insert(cboConfigure.Items.Count - 1, conf);
                cboConfigure.SelectedIndex = cboConfigure.Items.Count - 2;
            }
        }


    }
}