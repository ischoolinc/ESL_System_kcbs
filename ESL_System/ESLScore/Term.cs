using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System
{
    public class Term
    {
        /// <summary>
        /// 試別名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 試別分數比例
        /// </summary>
        public string Weight { get; set; }

        /// <summary>
        /// 試別分數輸入開始時間
        /// </summary>
        public string InputStartTime { get; set; }

        /// <summary>
        /// 試別分數輸入結束時間
        /// </summary>
        public string InputEndTime { get; set; }

        /// <summary>
        /// 自訂義子項目分數輸入開始時間
        /// </summary>
        public string CustomInputStartTime { get; set; }

        /// <summary>
        /// 自訂義子項目分數輸入結束時間
        /// </summary>
        public string CustomInputEndTime { get; set; }

        /// <summary>
        /// 參考試別id
        /// </summary>
        public string Ref_exam_id { get; set; }

        /// <summary>
        /// 子項目List
        /// </summary>
        public List<Subject> SubjectList { get; set; }

        /// <summary>
        /// 子項目總比例
        /// </summary>
        public decimal SubjectTotalWeight { get; set; }





    }
}
