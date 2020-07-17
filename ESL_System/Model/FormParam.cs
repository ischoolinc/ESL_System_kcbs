using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System.Model
{

    /// <summary>
    /// 傳遞From畫面資訊
    /// </summary>
    class FormParam
    {
        public FormParam(string SchoolYear, string semeter)
        {
            int _schoolYear;
            int _semester;

            Boolean isYearScuccess = Int32.TryParse(SchoolYear, out _schoolYear);
            Boolean isSemesterSucces = Int32.TryParse(semeter, out _semester);

            if (isYearScuccess)
            {
                this.SchoolYear = _schoolYear;
            }
            if (isSemesterSucces)
            {
                this.Semester = _semester;
            }

        }
        /// <summary>
        /// 學年度 
        /// </summary>
        public int SchoolYear { get; set; }

        /// <summary>
        /// 學期
        /// </summary>
        public int Semester { get; set; }






    }
}
