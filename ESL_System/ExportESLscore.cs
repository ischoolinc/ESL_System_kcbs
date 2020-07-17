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

namespace ESL_System
{
    //2018/11/16 穎驊 新增 產出 ESL 成績清單
    class ExportESLscore
    {
        private BackgroundWorker _worker;
        private List<string> _courseIDList;        

        public ExportESLscore(List<string> courseIDList)
        {
            _courseIDList = courseIDList;
        }

        public void export()
        {
            _worker = new BackgroundWorker();
            _worker.DoWork += new DoWorkEventHandler(Worker_DoWork);
            _worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Worker_RunWorkerCompleted);
            _worker.ProgressChanged += new ProgressChangedEventHandler(Worker_ProgressChanged);
            _worker.WorkerReportsProgress = true;

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
                }
            }
            #endregion

            if (_courseIDList.Count == 0)
            {
                return; // 沒有任何ESL: 課程，結束。
            }

            string courseeESLIDs = string.Join(",", _courseIDList); // 真正是ESL 課程的ID

            _worker.ReportProgress(50, "取得成績資料...");

            string scoreQuery = @"
WITH RawData AS 
(
    SELECT 
        $esl.gradebook_assessment_score.uid
        ,student.student_number
        ,student.name
        ,student.english_name
        ,class.class_name
        ,student.seat_no
        ,course.course_name
        ,teacher.teacher_name   
        ,$esl.gradebook_assessment_score.term
        ,$esl.gradebook_assessment_score.subject
        ,$esl.gradebook_assessment_score.assessment 
        ,$esl.gradebook_assessment_score.value
        ,CASE
    WHEN  $esl.gradebook_assessment_score.term IS NULL THEN 0
    WHEN  $esl.gradebook_assessment_score.term ='' THEN 0
    ELSE 1
    END
    ""term_has_value""
    ,CASE
    WHEN  $esl.gradebook_assessment_score.subject IS NULL THEN 0
    WHEN  $esl.gradebook_assessment_score.subject ='' THEN 0
    ELSE 1
    END
    ""subject_has_value""
    ,CASE
    WHEN  $esl.gradebook_assessment_score.assessment IS NULL THEN 0
    WHEN  $esl.gradebook_assessment_score.assessment ='' THEN 0
    ELSE 1
    END
    ""assessment_has_value""
    FROM $esl.gradebook_assessment_score
    LEFT JOIN student
    ON student.id = $esl.gradebook_assessment_score.ref_student_id
    LEFT JOIN class
    ON class.id = student.ref_class_id
    LEFT JOIN course
    ON course.id =  $esl.gradebook_assessment_score.ref_course_id
    LEFT JOIN teacher
    ON teacher.id =  $esl.gradebook_assessment_score.ref_teacher_id
    WHERE course.id IN( " + courseIDs + @")
    AND ($esl.gradebook_assessment_score.custom_assessment IS NULL OR $esl.gradebook_assessment_score.custom_assessment ='')    
)
SELECT
RawData.student_number
, RawData.name
, RawData.english_name
, RawData.class_name
, RawData.seat_no
, RawData.course_name
, RawData.teacher_name
, RawData.term
, RawData.subject
, RawData.assessment
, CASE RawData.assessment_has_value +RawData.subject_has_value +RawData.term_has_value
WHEN   1 THEN 'term 分數'
WHEN   2 THEN 'subject 分數'
WHEN   3 THEN 'assessment 分數'
    ELSE ''
    END
    ""score_type""
, RawData.value
FROM RawData
WHERE value !=''
ORDER BY course_name, student_number, term, subject, assessment
                      ";

            // 取的分數資料的表            
            DataTable scoreDT = qh.Select(scoreQuery);

            Workbook book = new Workbook();
            book.Worksheets.Clear();
            Worksheet ws = book.Worksheets[book.Worksheets.Add()];
            ws.Name = "ESL課程成績匯出";

            List<string> colheaderList = new List<string>();

            colheaderList.Add("student_number");
            colheaderList.Add("name");
            colheaderList.Add("english_name");
            colheaderList.Add("class_name");
            colheaderList.Add("seat_no");
            colheaderList.Add("course_name");
            colheaderList.Add("teacher_name");
            colheaderList.Add("term");
            colheaderList.Add("subject");
            colheaderList.Add("assessment");
            colheaderList.Add("score_type");
            colheaderList.Add("value");

            int columnIndex = 0;

            // 加入表頭
            foreach (string header in colheaderList)
            {                
                ws.Cells[0, columnIndex].PutValue(header);
                columnIndex++;
            }

            _worker.ReportProgress(80, "產生Excel報表...");

            //整理目前的ESL 課程資料
            if (scoreDT.Rows.Count > 0)
            {
                int rowIndex = 1; //0為表頭，這裡從1 開始

                foreach (DataRow dr in scoreDT.Rows)
                {
                    #region 填入內容

                    columnIndex = 0;

                    foreach (string header in colheaderList)
                    {
                        ws.Cells[rowIndex, columnIndex].PutValue("" + dr[header]);
                        columnIndex++;
                    }

                    rowIndex++;

                    _worker.ReportProgress(80+((20* rowIndex )/ scoreDT.Rows.Count), "產生Excel報表...");

                    #endregion
                }
            }

            ws.AutoFitColumns(); // 使 匯出excel 自動調整 欄寬

            e.Result = book;

            _worker.ReportProgress(100, "ESL課程成績匯出報表，產生完成。");
        }


        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Workbook book = new Workbook();

            if (e.Result == null)
            {
                MsgBox.Show("本次選擇範圍沒有套用ESL樣版的課程!");
                return;
            }
            else
            {
                book = (Workbook)e.Result;
            }
            
            SaveFileDialog sd = new SaveFileDialog();
            sd.FileName = "ESL課程成績匯出";
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

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage(""+e.UserState, e.ProgressPercentage);
        }
    }
}
