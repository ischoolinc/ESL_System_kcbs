using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System
{
    public class Subject
    {
        /// <summary>
        /// 子項目名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 子項目分數比例
        /// </summary>
        public string Weight { get; set; }

        /// <summary>
        /// 子評量List
        /// </summary>
        public List<Assessment> AssessmentList { get; set; }

        /// <summary>
        /// 子項目總比例
        /// </summary>
        public decimal AssessmentTotalWeight { get; set; }


    }
}
