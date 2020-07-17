using ESL_System.Model;
using ESL_System.UDT;
using FISCA.Data;
using FISCA.UDT;
using K12.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System.Service
{
    /// <summary>
    ///  撈取資料 整理資料用
    /// </summary>
    class DataService
    {
        private QueryHelper QHlper = new QueryHelper();
        List<SubjectInfoForGAPLevel> SubGPAMapping;
        List<ScoreGPAMapping> ScoreGPAMapping;

        public DataService()
        {
            AccessHelper accessHelper = new AccessHelper();
            SubGPAMapping = accessHelper.Select<SubjectInfoForGAPLevel>();
            ScoreGPAMapping = accessHelper.Select<ScoreGPAMapping>();
        }

        /// <summary>
        /// 取得學生學期科目文字描述
        /// </summary>
        /// <param name="studentIDs"></param>
        /// <param name="schoolYear"></param>
        /// <param name="semester"></param>
        /// <returns></returns>
        public  Dictionary<string, Dictionary<string, string>> GetSemsSubjText(List<string> studentIDs, int schoolYear, int semester)
        {
            Dictionary<string, Dictionary<string, string>> dicSubjTextInfos = new Dictionary<string, Dictionary<string, string>>();

            string sql = @"
SELECT
	sems_subj_score_ext.ref_student_id
	, sems_subj_score_ext.grade_year
	, sems_subj_score_ext.semester
	, sems_subj_score_ext.school_year
	, array_to_string(xpath('//Subject/@科目', subj_score_ele), '')::text AS subject 
    , array_to_string(xpath('//Subject/@文字描述', subj_score_ele), '')::text AS 文字描述
FROM (
		SELECT 
			sems_subj_score.*
			, 	unnest(xpath('//SemesterSubjectScoreInfo/Subject', xmlparse(content  '<root>' ||score_info||'</root>' ))) as subj_score_ele
		FROM 
			sems_subj_score "
    + $"WHERE ref_student_id  IN ( {string.Join(",",studentIDs)}) AND school_year ={schoolYear} AND semester={semester}) as sems_subj_score_ext";

            DataTable dt = this.QHlper.Select(sql);

            foreach (DataRow dr in  dt.Rows) 
            {
                string ref_student_id = "" + dr["ref_student_id"];
                string SubjectName = "" + dr["subject"];
                string SemsScoreText = "" + dr["文字描述"];

                if (!dicSubjTextInfos.ContainsKey(ref_student_id))
                {
                    dicSubjTextInfos.Add(ref_student_id, new Dictionary<string, string>());

                }

                if (!dicSubjTextInfos[ref_student_id].ContainsKey(SubjectName))
                {
                    dicSubjTextInfos[ref_student_id].Add(SubjectName, SemsScoreText);
                }


            }

        
            return dicSubjTextInfos;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public decimal? GetFinalGPA(string subject, decimal score)
        {
            bool Iscontain = this.SubGPAMapping.Any(x => x.Subject == subject);
            if (Iscontain)
            {
                SubjectInfoForGAPLevel maping = this.SubGPAMapping.Where(x => x.Subject == subject).First();

                if (maping.IsAP)
                {
                    return GetAP(score);
                }
                else if (maping.IsHoner)
                {
                    return GetHoner(score);
                }
                else  // standandard
                {
                    return GetStandar(score);
                }
            }
            else
            {
                return GetStandar(score);
            }

        }



        public decimal? GetHoner(decimal score)
        {
            foreach (ScoreGPAMapping scoreGPAMapping in this.ScoreGPAMapping)
            {
                if (score != 100)
                {
                    if (score >= scoreGPAMapping.MinScore && score < scoreGPAMapping.MaxScore)
                    {
                        return scoreGPAMapping.Honers;
                    }
                }
                else
                {
                    if (score >= scoreGPAMapping.MinScore && score <= scoreGPAMapping.MaxScore)
                    {
                        return scoreGPAMapping.Honers;
                    }
                }
            }
            return null;
        }



        public decimal? GetAP(decimal score)
        {
            foreach (ScoreGPAMapping scoreGPAMapping in this.ScoreGPAMapping)
            {
                if (score != 100)
                {
                    if (score >= scoreGPAMapping.MinScore && score < scoreGPAMapping.MaxScore)
                    {
                        return scoreGPAMapping.AP;
                    }
                }
                else  // 分數等於100 
                {
                    if (score >= scoreGPAMapping.MinScore && score <= scoreGPAMapping.MaxScore)
                    {
                        return scoreGPAMapping.AP;
                    }
                }

            }
            return null;

        }


        public decimal? GetStandar(decimal score)
        {
            foreach (ScoreGPAMapping scoreGPAMapping in this.ScoreGPAMapping)
            {
                if (score != 100)
                {
                    if (score >= scoreGPAMapping.MinScore && score < scoreGPAMapping.MaxScore)
                    {
                        return scoreGPAMapping.GPA;

                    }
                }
                else  //分數=100
                {

                    if (score >= scoreGPAMapping.MinScore && score <= scoreGPAMapping.MaxScore)
                    {
                        return scoreGPAMapping.GPA;

                    }
                }

            }
            return null;

        }


        /// <summary>
        /// 取得領域清單
        /// </summary>
        /// <returns></returns>

        public Dictionary<string, int> GetDomainOrderDic() 
        {
            Dictionary<string, int> domainOrder =new Dictionary<string, int>();
            string sql = @"  SELECT
   
	 unnest(xpath('//Domains/Domain/@Name',  xmlparse(content replace(replace(content ,'&lt;','<'),'&gt;','>'))))::text AS domain_name
	FROM  
		list 
	WHERE name  ='JHEvaluation_Subject_Ordinal'";
           DataTable dt= this.QHlper.Select(sql);
            int order = 1;
            foreach (DataRow dr  in dt.Rows)
            {
                domainOrder.Add(""+dr["domain_name"],order++);
            }


            return domainOrder;

        }


        /// <summary>
        /// 取得學生修課紀錄整理 科目排序是照領域排
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, StudentScAttend> GetStudentScAttend(int schoolYear, int semester, List<string> courseIDs, List<string> studentIDs)
        {
            //載入領域
            Dictionary<string, int> DomainOrder = this.GetDomainOrderDic();


            Dictionary<string, StudentScAttend> dicStudentScAttend = new Dictionary<string, StudentScAttend>();

            string sql = @"
SELECT
     course.school_year
    , course.semester
    , sc_attend.ref_student_id 
	, course.id 
	, course.course_name 
	, course.subject
    , course.domain
FROM 
	(SELECT *FROM 	course	WHERE   id IN   ({0})) AS course 
INNER  JOIN
	(SELECT * FROM  sc_attend  WHERE ref_student_id IN ({1})) AS sc_attend 
ON
    sc_attend.ref_course_id = course.id 
ORDER BY  domain ,subject";
            sql = string.Format(sql, string.Join(",", courseIDs), string.Join(",", studentIDs));


            DataTable dt = this.QHlper.Select(sql);


            foreach (DataRow dr in dt.Rows)
            {

                string studentID = "" + dr["ref_student_id"];
                string subject = "" + dr["subject"];
                string domain = "" + dr["domain"];
                int domainOrder = DomainOrder.ContainsKey(domain) ? DomainOrder[domain] : 99999;

                if (!dicStudentScAttend.ContainsKey(studentID)) // 如果還沒有此學生之資料就在dic裡增加學生的空間
                {
                    dicStudentScAttend.Add(studentID, new StudentScAttend(studentID, schoolYear, semester));
                }

                if (!dicStudentScAttend[studentID].SubjectFromScAttend.Any(x=>x.Subject == subject)) // 如果此學生不包含此科目
                {
                    dicStudentScAttend[studentID].SubjectFromScAttend.Add(new CourseInfo(subject, domain, domainOrder) ); // 把修課紀錄加進去
                }

            }



            //排序每個學生之領域



            return dicStudentScAttend;
        }
    }
}
