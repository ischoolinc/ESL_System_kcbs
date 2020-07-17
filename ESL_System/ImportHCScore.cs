using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using FISCA.Presentation.Controls;
using Aspose.Cells;
using K12.Data;

namespace ESL_System
{

    // 2018/11/21 穎驊備註，本功能為學校有了自行打分數的EXCEL後，工程人員協助產生匯入ESL 成績的SQL 文字，
    // 再使用ischool 中央系統，一次將成績匯入， EXCEL 的格式可見 Resource 內的範例。
    class ImportHCScore
    {

        List<ESLScore> insertESLscoreList = new List<ESLScore>();

        public ImportHCScore()
        {
            OpenFileDialog ope = new OpenFileDialog();
            ope.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            Workbook wb = new Workbook();

            if (ope.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            else
            {
                wb = new Workbook(ope.FileName);
            }

            Worksheet ws = wb.Worksheets[0];

            Cells cells = ws.Cells;



            foreach (var row in cells.Rows)
            {
                if (row.Index > 1 && !row.IsBlank)
                {
                    //LA CET 成績
                    for (int i = 16; i <= 22; i++)
                    {
                        ESLScore score = new ESLScore();

                        score.RefCourseID = "" + cells[row.Index, 3].Value;

                        score.RefStudentID = "" + cells[row.Index, 7].Value;

                        score.RefTeacherID = "" + cells[row.Index, 13].Value;

                        score.Term = "Mid-Term";

                        score.Subject = "Language Arts (CET)";

                        score.Assessment = "" + cells[1, i].Value;

                        score.Value = "" + cells[row.Index, i].Value;

                        this.insertESLscoreList.Add(score);
                    }

                    //LA CET 成績
                    for (int i = 23; i <= 29; i++)
                    {
                        ESLScore score = new ESLScore();

                        score.RefCourseID = "" + cells[row.Index, 3].Value;

                        score.RefStudentID = "" + cells[row.Index, 7].Value;

                        score.RefTeacherID = "" + cells[row.Index, 14].Value;

                        score.Term = "Mid-Term";

                        score.Subject = "Language Arts (FET)";

                        score.Assessment = "" + cells[1, i].Value;

                        score.Value = "" + cells[row.Index, i].Value;

                        this.insertESLscoreList.Add(score);
                    }

                    //SC
                    for (int i = 30; i <= 35; i++)
                    {
                        ESLScore score = new ESLScore();

                        score.RefCourseID = "" + cells[row.Index, 5].Value;

                        score.RefStudentID = "" + cells[row.Index, 7].Value;

                        score.RefTeacherID = "" + cells[row.Index, 15].Value;

                        score.Term = "Mid-Term";

                        score.Subject = "Science";

                        score.Assessment = "" + cells[1, i].Value;

                        score.Value = "" + cells[row.Index, i].Value;

                        this.insertESLscoreList.Add(score);
                    }
                }
            }

            //拚SQL
            // 兜資料
            List<string> dataList = new List<string>();

            

            foreach (ESLScore score in insertESLscoreList)
            {
                string data = string.Format(@"
                SELECT
                    '{0}'::BIGINT AS ref_student_id
                    ,'{1}'::BIGINT AS ref_course_id
                    ,'{2}'::BIGINT AS ref_teacher_id
                    ,'{3}'::TEXT AS term
                    ,{4} AS subject
                    ,'{5}'::TEXT AS assessment
                    ,'{6}'::TEXT AS value
                    ,{7}::INTEGER AS uid
                    ,'INSERT'::TEXT AS action
                ", score.RefStudentID, score.RefCourseID, score.RefTeacherID, score.Term, score.Subject != null ? "'" + score.Subject + "' ::TEXT" : "NULL",score.Assessment, score.Value, 0);  // insert 給 uid = 0

                dataList.Add(data);
            }

            string Data = string.Join(" UNION ALL", dataList);


            string sql = string.Format(@"
WITH score_data_row AS(			 
                {0}     
)
INSERT INTO $esl.gradebook_assessment_score(
	ref_student_id	
	,ref_course_id
    ,ref_teacher_id
	,term
	,subject
    ,assessment
	,value
)
SELECT 
	score_data_row.ref_student_id::BIGINT AS ref_student_id	
	,score_data_row.ref_course_id::BIGINT AS ref_course_id
    ,score_data_row.ref_teacher_id::BIGINT AS ref_teacher_id
	,score_data_row.term::TEXT AS term	
	,score_data_row.subject::TEXT AS subject	
    ,score_data_row.assessment::TEXT AS subject
	,score_data_row.value::TEXT AS value	
FROM
	score_data_row
WHERE action ='INSERT'", Data);


        }




    }
}


