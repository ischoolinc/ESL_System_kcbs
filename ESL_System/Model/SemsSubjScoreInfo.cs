using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System.Model
{

    /// <summary>
    /// 
    /// </summary>
    class SemsSubjScoreInfo
    {
        public SemsSubjScoreInfo(string subject ,decimal semsScore, decimal? semsGPA ) 
        {
            this.Subject = subject;
            this.SemsScore = semsScore;
            this.SemsGPA = semsGPA;
        }


        /// <summary>
        /// 科目名稱
        /// </summary>
        public string  Subject { get; set; }

        /// <summary>
        /// 科目學期成績
        /// </summary>
        public decimal? SemsScore { get; set; }


        /// <summary>
        /// 科目學期成績GPA
        /// </summary>
        public decimal? SemsGPA { get; set; }

    }
}
