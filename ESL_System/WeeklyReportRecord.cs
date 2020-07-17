using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using FISCA.Data;
using K12.Data;

namespace ESL_System
{
    class ESLCourseRecord
    {
        //
        // Summary:
        //     課程授課教師列表
        public List<CourseTeacherRecord> ESLTeachers { get; set; }
        //
        // Summary:
        //     主要授課教師暱稱
        [Field(Caption = "教師暱稱（主要授課）", EntityName = "Teacher", EntityCaption = "教師")]
        public string ESLMajorTeacherNickname { get; set; }
        //
        // Summary:
        //     主要授課教師名稱
        [Field(Caption = "教師名稱（主要授課）", EntityName = "Teacher", EntityCaption = "教師")]
        public string ESLMajorTeacherName { get; set; }
        //
        // Summary:
        //     主要授課教師編號
        [Field(Caption = "教師編號（主要授課）", EntityName = "Teacher", EntityCaption = "教師", IsEntityPrimaryKey = true)]
        public string ESLMajorTeacherID { get; set; }
        //
        // Summary:
        //     所屬試別設定
        public AssessmentSetupRecord ESLAssessmentSetup { get; set; }
        //
        // Summary:
        //     所屬班級
        public ClassRecord ESLClass { get; set; }
        //
        // Summary:
        //     所屬班級編號
        [Field(Caption = "班級編號", EntityName = "Class", EntityCaption = "班級", IsEntityPrimaryKey = true)]
        public string ESLRefClassID { get; set; }
        //
        // Summary:
        //     權數，相當於高中的學分數
        [Field(Caption = "權數(學分數)", EntityName = "Course", EntityCaption = "課程")]
        public decimal? ESLCredit { get; set; }
        //
        // Summary:
        //     所屬試別設定編號
        [Field(Caption = "評量設定編號", EntityName = "AssessmentSetup", EntityCaption = "評量設定", IsEntityPrimaryKey = true)]
        public string ESLRefAssessmentSetupID { get; set; }
        //
        // Summary:
        //     科目
        [Field(Caption = "科目", EntityName = "Course", EntityCaption = "課程")]
        public string ESLSubject { get; set; }
        //
        // Summary:
        //     系統編號
        [Field(Caption = "編號", EntityName = "Course", EntityCaption = "課程", IsEntityPrimaryKey = true)]
        public string ESLID { get; set; }
        //
        // Summary:
        //     名稱
        [Field(Caption = "名稱", EntityName = "Course", EntityCaption = "課程")]
        public string ESLName { get; set; }
        //
        // Summary:
        //     學年度
        [Field(Caption = "學年度", EntityName = "Course", EntityCaption = "課程")]
        public int? ESLSchoolYear { get; set; }
        //
        // Summary:
        //     學期
        [Field(Caption = "學期", EntityName = "Course", EntityCaption = "課程")]
        public int? ESLSemester { get; set; }
        //
        // Summary:
        //     節數，實際的上課時數
        [Field(Caption = "節數", EntityName = "Course", EntityCaption = "課程")]
        public decimal? ESLPeriod { get; set; }
        //
        // Summary:
        //     課程難度(Level)
        [Field(Caption = "課程難度(Level)", EntityName = "Course", EntityCaption = "課程")]
        public string ESLDifficulty { get; set; }

        public static List<ESLCourseRecord> ToESLCourseRecords(List<K12.Data.CourseRecord> courseList)
        {
            List<ESLCourseRecord> eslCourseList = new List<ESLCourseRecord>();
            string courseIDs = string.Join(",", courseList.Select(x => x.ID).ToList());
            string selectSQL = @"
SELECT
    id
    , difficulty
FROM
    course
WHERE
    id in 
    (
        " + courseIDs + @"
    )
";
            QueryHelper queryHelper = new QueryHelper();
            DataTable dt = queryHelper.Select(selectSQL);
            Dictionary<string, string> courseDic = new Dictionary<string, string>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (!courseDic.ContainsKey("" + dt.Rows[i]["id"]))
                {
                    courseDic.Add("" + dt.Rows[i]["id"], null);
                }
                courseDic["" + dt.Rows[i]["id"]] = "" + dt.Rows[i]["difficulty"];
            }

            foreach (CourseRecord courseRecord in courseList)
            {
                ESLCourseRecord eslCourse = new ESLCourseRecord();

                //Teachers
                eslCourse.ESLTeachers = courseRecord.Teachers;

                //MajorTeacherNickname
                eslCourse.ESLMajorTeacherNickname = courseRecord.MajorTeacherNickname;

                //MajorTeacherName
                eslCourse.ESLMajorTeacherName = courseRecord.MajorTeacherName;

                //MajorTeacherID
                eslCourse.ESLMajorTeacherID = courseRecord.MajorTeacherID;

                //AssessmentSetup
                eslCourse.ESLAssessmentSetup = courseRecord.AssessmentSetup;

                //Class
                eslCourse.ESLClass = courseRecord.Class;

                //RefClassID
                eslCourse.ESLRefClassID = courseRecord.RefClassID;

                //Credit
                eslCourse.ESLCredit = courseRecord.Credit;

                //RefAssessmentSetupID
                eslCourse.ESLRefAssessmentSetupID = courseRecord.RefAssessmentSetupID;

                //Subject
                eslCourse.ESLSubject = courseRecord.Subject;

                //ID
                eslCourse.ESLID = courseRecord.ID;

                //Name
                eslCourse.ESLName = courseRecord.Name;

                //SchoolYear
                eslCourse.ESLSchoolYear = courseRecord.SchoolYear;

                //Semester
                eslCourse.ESLSemester = courseRecord.Semester;

                //Period
                eslCourse.ESLPeriod = courseRecord.Period;

                //Diffiiculty
                eslCourse.ESLDifficulty = courseDic[courseRecord.ID];

                eslCourseList.Add(eslCourse);
            }

            return eslCourseList;
        }
    }
}
