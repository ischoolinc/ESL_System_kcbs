using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System.Model
{

    //20200714 為了讓定期評量之科目與學期成績之科目都能夠印出來
    class StudentScAttend
    {
        /// <summary>
        ///  建構子
        /// </summary>
        public StudentScAttend(string studentID ,int schoolYear ,int semester )
        {
            this.RefStudentID = studentID;
            this.SchoolYear = schoolYear;
            this.Semester = semester;
            this.SubjectFromScAttend = new List<CourseInfo>();
        }
        /// <summary>
        /// 學生系統ID 
        /// </summary>
        public string RefStudentID { get; set; }

        /// <summary>
        /// 學年度
        /// </summary>
        public int SchoolYear { get; set; }

        /// <summary>
        /// 學期
        /// </summary>
        public int Semester { get; set; }

        /// <summary>
        /// 該學生當學年度學期修課紀錄上修過之科目
        /// </summary>
        public List<CourseInfo> SubjectFromScAttend { get; set; }
    }
}
