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

namespace ESL_System.Form
{
    public partial class PrintWeeklyReportForm : FISCA.Presentation.Controls.BaseForm
    {
        private BackgroundWorker _bw = new BackgroundWorker();

        private bool _isOrderBySubject = false;
        private string _orderedSubject = ""; // 排序科目

        private string school_year = "";
        private string semester = "";

        //// 儲放weeklyReport 學生 behaviorItem 資料 的dict <studentID,<behaviorItemKey,behaviorItemValue>
        //private Dictionary<string, Dictionary<string, string>> _behaviorItemDict = new Dictionary<string, Dictionary<string, string>>();

        //// 儲放weeklyReport 學生 gradebook 資料 的dict <studentID,<scoreItemKey,scoreItemValue>
        //private Dictionary<string, Dictionary<string, string>> _weeklyScoreDict = new Dictionary<string, Dictionary<string, string>>();

        //// [ischoolkingdom] Vicky新增，WeeklyReport報表列印，新增可輸出 課程名稱、教師名稱
        //// 儲放weeklyReport 學生 其他 資料 的dict
        //private Dictionary<string, Dictionary<string, string>> _otherItemDict = new Dictionary<string, Dictionary<string, string>>();

        // 儲放weeklyReport 學生 gradebook 資料 的dict <studentID,<ItemKey,ItemValue>
        private Dictionary<string, List<WeeklyReportRecord>> _weeklyDataDict = new Dictionary<string, List<WeeklyReportRecord>>();


        // 儲放 所有課程的科目名稱
        private List<string> _subjectList = new List<string>();

        BackgroundWorker bkw;

        // 開始日期
        private DateTime _BeginDate;
        // 結束日期
        private DateTime _EndDate;
        
        private DataTable _mergeDataTable = new DataTable();

        private List<UDT_WeeklyReportTemplate> _configuresList = new List<UDT_WeeklyReportTemplate>();

        private UDT_WeeklyReportTemplate _configure { get; set; }

        public PrintWeeklyReportForm()
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
            
            // [ischoolkingdom] Vicky新增，WeeklyReport報表列印，預設時間為當週
            // 預設都為今天之當週週一至週五   
            DateTime date = DateTime.Now;
            if (date.DayOfWeek != DayOfWeek.Monday)
            {
                if (date.DayOfWeek == DayOfWeek.Sunday)
                {
                    date.AddDays(-1);
                }

                DateTime modified_date = date.AddDays(Convert.ToDouble(1 - (int)date.DayOfWeek)); //把date調整至週一
                dtBegin.Value = modified_date;
                dtEnd.Value = modified_date.AddDays(4);
            }
            else
            {
                dtBegin.Value = date;
                dtEnd.Value = date.AddDays(4);
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


            bkw.ReportProgress(80);

            FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

            //string qry = "TemplateStream IS NOT null";

            _configuresList = _AccessHelper.Select<UDT_WeeklyReportTemplate>();

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
            cboConfigure.Items.Add(new UDT_WeeklyReportTemplate() { Name = "新增" });

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

            school_year = comboBoxEx1.Text;
            semester = comboBoxEx2.Text;

            // 假如使用者有選擇科目，列印順序將以 課程科目排序、而非班級學生
            if (comboBoxEx3.Text != "")
            {
                _isOrderBySubject = true;
                _orderedSubject = comboBoxEx3.Text; // 排序科目
            }

            _bw.RunWorkerAsync();

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

            string reportName = "WeeklyReport樣板(" + _configure.Name + ")";

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

            #region WeeklyReport報表統計

            builder.ParagraphFormat.Style = builder.Document.Styles["ESLNameStyle"];
            // 固定變數，不分　期中、期末、學期  (使用大字粗體)
            builder.Writeln("WeeklyReport報表統計");

            builder.ParagraphFormat.Style = builder.Document.Styles["Normal"];

            builder.StartTable();
            builder.InsertCell();
            builder.Write("項目");
            builder.InsertCell();
            builder.Write("變數");
            builder.EndRow();

            List<string> itemList = new List<string>();
            // [ischoolkingdom] Vicky新增，WeeklyReport報表列印，課程名稱、教師名稱功能變數
            itemList.Add("課程名稱");
            itemList.Add("教師名稱");
            itemList.Add("區間開始日期");
            itemList.Add("區間結束日期");
            itemList.Add("Performance資料");
            itemList.Add("Score資料");
            itemList.Add("GeneralComment");
            itemList.Add("PersonalComment");

            foreach (string key in itemList)
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
            #region 成績變數


            // 2018/6/15 穎驊備註 以下整理 功能變數 最常使用的 string..Trim().Replace(' ', '_').Replace('"', '_') 
            // >> 其用意為避免Word 功能變數合併列印時 會有一些奇怪的BUG ，EX: row["Final-Term評量_Science科目_In-Class Score子項目_分數1"] = "YOYO!"; >> 有空格印不出來 


            // Apply the paragraph style to the current paragraph in the document and add some text.
            builder.ParagraphFormat.Style = builder.Document.Styles["ESLNameStyle"];
            // 每一個 評量的名稱 放在最上面 (使用大字粗體)
            builder.Writeln("評量名稱: " + ExamName);

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
                //builder.InsertCell();
                //builder.Write("科目文字評量");
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

                    builder.EndRow();
                }

                builder.EndTable();
                builder.Writeln();
            }

            #endregion
        }

        private void _bkWork_DoWork(object sender, DoWorkEventArgs e)
        {
            // 處理等第
            DegreeMapper dm = new DegreeMapper();

            List<string> studentIDList = new List<string>();
            List<string> courseIDList = new List<string>();

            // 選擇的學生名單 
            studentIDList = K12.Presentation.NLDPanels.Student.SelectedSource;


            #region 取得本學期有設定ESL 評分樣版的課程清單 (才會有weeklyReport 的資料)
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
                MsgBox.Show("本學期沒有任何設定ESL樣板的課程。", "錯誤!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion


            //取得學生基本資料
            List<K12.Data.StudentRecord> studentList = K12.Data.Student.SelectByIDs(studentIDList);


            #region 取得Weekly 資料
            _bw.ReportProgress(20, "取得ESL課程WeeklyReport 資料");

            string course_ids = string.Join("','", courseIDList);

            string student_ids = string.Join("','", studentIDList);

            
            // 建立 weeklyScore結構
            foreach (string stuID in studentIDList)
            {
                _weeklyDataDict.Add(stuID, new List<WeeklyReportRecord>());
            }


            // [ischoolkingdom] Vicky新增，WeeklyReport報表列印，sql加入 課程名稱、教師名稱 資料
            // 按照時間順序抓WeeklyReport
            string sqlWeeklyReport = @"
SELECT 
	$esl.weekly_report.ref_course_id AS ref_course_id
	,$esl.weekly_report.ref_teacher_id AS ref_teacher_id
	,$esl.weekly_data.ref_student_id AS ref_student_id
	,$esl.weekly_report.begin_date
	,$esl.weekly_report.end_date
	,$esl.weekly_data.grade_book_data
	,$esl.weekly_data.behavior_data	
	,$esl.weekly_report.general_comment 	
	,$esl.weekly_data.personal_comment
    ,teacher.teacher_name
    ,course.course_name
    ,course.subject
FROM $esl.weekly_report
LEFT JOIN teacher ON $esl.weekly_report.ref_teacher_id = teacher.id
LEFT JOIN course ON $esl.weekly_report.ref_course_id = course.id
LEFT JOIN $esl.weekly_data  ON $esl.weekly_data.ref_weekly_report_uid = $esl.weekly_report.uid
WHERE 
'" + _BeginDate.ToString("yyyy/MM/dd") + " 00:00:00'" + @" <=$esl.weekly_report.begin_date
AND '" + _EndDate.ToString("yyyy/MM/dd") + " 23:59:59'" + @" >=$esl.weekly_report.begin_date
AND ref_course_id IN ('" + course_ids + @"') 
AND ref_student_id IN('" + student_ids + @"')
ORDER BY ref_student_id ";

            DataTable dtWeeklyData = qh.Select(sqlWeeklyReport);

            decimal progress = 20;
            decimal per = (decimal)(100 - progress) / (dtWeeklyData.Rows.Count != 0 ? dtWeeklyData.Rows.Count : 1);



            foreach (DataRow row in dtWeeklyData.Rows)
            {
                string id = "" + row["ref_student_id"];

                string grade_book_data = "" + row["grade_book_data"];

                string behavior_data = "" + row["behavior_data"];

                string general_comment = "" + row["general_comment"];

                string personal_comment = "" + row["personal_comment"];

                // [ischoolkingdom] Vicky新增，WeeklyReport報表列印，課程名稱、教師名稱功能變數
                string course_name = "" + row["course_name"];

                string teacher_name = "" + row["teacher_name"];

                string subject = "" + row["subject"];

                WeeklyReportRecord wr = new WeeklyReportRecord();

                if (_weeklyDataDict.ContainsKey(id))
                {
                    {
                        string modified_text = grade_book_data.Insert(0, "{\"data\":");
                        string modified_text_data = modified_text.Insert(modified_text.Length, "}");
                        Json_deserialize_data json_grade_data = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<Json_deserialize_data>(modified_text_data);
                        string k = "";
                        foreach (var item in json_grade_data.data)
                        {
                            List<string> gradelist = new List<string>();
                            gradelist.Add(item.customAssessment + ": " + item.value);

                            foreach (var i in gradelist)
                            {
                                k += i + "\r\n";
                            }
                        }

                        wr.ScoreData = k;
                    }

                    {
                        string modified_text = behavior_data.Insert(0, "{\"data\":");
                        string modified_text_data = modified_text.Insert(modified_text.Length, "}");
                        Json_deserialize_data json_behavior_data = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<Json_deserialize_data>(modified_text_data);
                        string k = "";
                        foreach (var item in json_behavior_data.data)
                        {
                            List<string> behaviorcommentlist = new List<string>();
                            behaviorcommentlist.Add(item.comment + "(" + item.createdate2 + ")");

                            foreach (var i in behaviorcommentlist)
                            {
                                k += i + "\r\n";
                            }
                        }

                        wr.PerformanceData = k;
                    }

                    wr.GeneralComment = general_comment;

                    wr.PersonalComment = personal_comment;

                    wr.CourseName = course_name;

                    wr.TeacherName = teacher_name;

                    // 排序用屬性，實際不印出
                    wr.Subject = subject;

                    _weeklyDataDict[id].Add(wr);
                }
            }

            #endregion

            // BY subject 排序 ，將studentList 順序重新整理
            if (_isOrderBySubject)
            {
                // 
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
            //  排序每一個學生 內的每一個課程
            foreach (StudentRecord stuRecord in studentList)
            {
                string id = stuRecord.ID;
                if (_weeklyDataDict.ContainsKey(id))
                {
                    List<WeeklyReportRecord> WeeklyDataList_new = new List<WeeklyReportRecord>();

                    WeeklyReportRecord wr = _weeklyDataDict[id].Find(x => x.Subject == _orderedSubject);

                    if (wr != null)
                    {
                        WeeklyDataList_new.Add(wr);

                        _weeklyDataDict[id].Remove(wr);
                    }

                    // 剩下的課程 用名稱排序
                    _weeklyDataDict[id].Sort((x, y) => { return x.CourseName.CompareTo(y.CourseName); });

                    WeeklyDataList_new.AddRange(_weeklyDataDict[id]);

                    _weeklyDataDict[id] = WeeklyDataList_new;
                }
            }

            // 將欄位 指派給 DataTable
            _mergeDataTable = GetMergeField();


            // Doc 切分字典(StuID,List<Doc>) 讓同一個學生若有多頁資料 可以是同一份電子報表上傳
            Dictionary<string, List<Document>> docDict = new Dictionary<string, List<Document>>();

            foreach (StudentRecord stuRecord in studentList)
            {
                string id = stuRecord.ID;

                Document _doc;

                // data book 資料
                if (_weeklyDataDict.ContainsKey(id))
                {
                    // 有項目 才列印
                    if (_weeklyDataDict[id].Count != 0)
                    {
                        foreach (WeeklyReportRecord wr in _weeklyDataDict[id])
                        {
                            DataRow row = _mergeDataTable.NewRow();

                            row["電子報表辨識編號"] = "系統編號{" + stuRecord.ID + "}"; // 學生系統編號

                            row["學年度"] = K12.Data.School.DefaultSchoolYear;
                            row["學期"] = K12.Data.School.DefaultSemester;
                            row["學號"] = stuRecord.StudentNumber;
                            row["年級"] = stuRecord.Class != null ? "" + stuRecord.Class.GradeYear : "";

                            row["原班級名稱"] = stuRecord.Class != null ? "" + stuRecord.Class.Name : "";
                            row["學生英文姓名"] = stuRecord.EnglishName;
                            row["學生中文姓名"] = stuRecord.Name;

                            row["區間開始日期"] = _BeginDate.ToShortDateString();
                            row["區間結束日期"] = _EndDate.ToShortDateString();


                            row["課程名稱"] = wr.CourseName;
                            row["教師名稱"] = wr.TeacherName;
                            row["Score資料"] = wr.ScoreData;
                            row["Performance資料"] = wr.PerformanceData;
                            row["GeneralComment"] = wr.GeneralComment;
                            row["PersonalComment"] = wr.PersonalComment;

                            _mergeDataTable.Rows.Add(row);

                        }

                        try
                        {
                            // 樣板的設定
                            _doc = _configure.Template.Clone();
                        }
                        catch (Exception ex)
                        {
                            MsgBox.Show(ex.Message);
                            e.Cancel = true;
                            return;
                        }


                        _doc.MailMerge.Execute(_mergeDataTable);

                        _mergeDataTable.Clear();

                        //for (int i = _doc.Sections.Count - 2; i >= 0; i--)
                        //{
                        //    // Copy the content of the current section to the beginning of the last section.
                        //    _doc.LastSection.PrependContent(_doc.Sections[i]);
                        //    // Remove the copied section.
                        //    _doc.Sections[i].Remove();
                        //}

                        if (!docDict.ContainsKey(stuRecord.ID))
                        {
                            docDict.Add(stuRecord.ID, new List<Document>());

                            docDict[stuRecord.ID].Add(_doc);
                        }
                        else
                        {
                            docDict[stuRecord.ID].Add(_doc);
                        }

                    }
                }                
            }





            //// Section 切分字典(StuID,List<Section>) 讓同一個學生若有多頁資料 可以是同一份電子報表上傳
            //Dictionary<string,Section> sectionDict = new Dictionary<string, Section>();

            Document doc_final = new Document();

            //foreach (Section each in _doc.Sections)
            //{
            //    Regex rx;
            //    Group g;

            //    rx = new Regex(@"系統編號\{([0-9a-zA-Z]+)\}");

            //    Match ch = rx.Match(each.GetText());
            //    if (ch.Success)
            //    {
            //        g = ch.Groups[1];

            //        if (!sectionDict.ContainsKey(g.Value))
            //        {
            //            sectionDict.Add(g.Value, new Section(doc_final));

            //            Node n = doc_final.ImportNode(each.FirstChild, true);

            //            sectionDict[g.Value].AppendChild(n);
            //        }
            //        else
            //        {
            //            Node n = doc_final.ImportNode(each.FirstChild, true);
            //            sectionDict[g.Value].AppendChild(n);
            //        }
            //    }
            //    else
            //    {
            //        if (!sectionDict.ContainsKey("沒有系統編號"))
            //        {
            //            sectionDict.Add("沒有系統編號", new Section(doc_final));

            //            Node n = doc_final.ImportNode(each.FirstChild, true);

            //            sectionDict["沒有系統編號"].AppendChild(n);
            //        }
            //        else
            //        {
            //            Node n = doc_final.ImportNode(each.FirstChild, true);
            //            sectionDict["沒有系統編號"].AppendChild(n);                        
            //        }
            //    }
            //}

            doc_final.Sections.Clear();

            foreach (string stuID in docDict.Keys)
            {
                foreach (Document d in docDict[stuID])
                {
                    for (int i = 0; i < d.Sections.Count; i++)
                    {
                        Node n = doc_final.ImportNode(d.Sections[i], true);
                        doc_final.Sections.Add(n);
                    }                    
                }                
            }

            e.Result = doc_final;
        }

        private void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                this.Close();
                return;
            }

            FISCA.Presentation.MotherForm.SetStatusBarMessage(" WeeklyReport報表產生完成。");

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
            sd.FileName = "WeeklyReport 報表.docx";
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

        private DataTable GetMergeField()
        {
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

            dataTable.Columns.Add("課程名稱");
            dataTable.Columns.Add("教師名稱");
            dataTable.Columns.Add("區間開始日期");
            dataTable.Columns.Add("區間結束日期");
            dataTable.Columns.Add("Performance資料");
            dataTable.Columns.Add("Score資料");
            dataTable.Columns.Add("GeneralComment");
            dataTable.Columns.Add("PersonalComment");

            #endregion

            return dataTable;
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
                    _configure = new UDT_WeeklyReportTemplate();
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
                UDT_WeeklyReportTemplate conf = new UDT_WeeklyReportTemplate();
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
    

    /// 寫data屬性
    public class Json_deserialize_data
    {
        public List<Json_data> data { get; set; }
    }
    public class Json_data
    {   //behavior
        public string createdate2 { get; set; }
        public string comment { get; set; }
        //score
        public string assessment { get; set; }
        public string customAssessment { get; set; }
        public string subject { get; set; }
        public string score_date { get; set; }
        public string description { get; set; }
        public string courseID { get; set; }
        public string teacherID { get; set; }
        public string value { get; set; }
    }


}
