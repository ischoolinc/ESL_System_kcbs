using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System
{
    // 本 物件 為仿 K12.Data.SCETakeRecord 暫時 在Local 端處理使用
    public class SCETakeESLRecord
    {

        /// <summary>
        /// 成績ID
        /// </summary>
        public string ID { get; set; }

        /// <summary>
        /// 參考課程ID
        /// </summary>
        public string RefCourseID { get; set; }

        /// <summary>
        /// 參考學生ID
        /// </summary>
        public string RefStudentID { get; set; }

        /// <summary>
        /// 參考學生修課編號ID
        /// </summary>
        public string RefSCAttendID { get; set; }

        /// <summary>
        /// 參考試別ID
        /// </summary>
        public string RefExamID { get; set; }


        /// <summary>
        /// 分數
        /// </summary>
        public decimal? Score { get; set; }


        /// <summary>
        /// 延伸資料 (國中的評量成績放在這裡)
        /// </summary>
        public string Extensions { get; set; }

        /// <summary>
        /// 資料需update
        /// </summary>
        public bool NeedUpdate  { get; set; }


    }
}
