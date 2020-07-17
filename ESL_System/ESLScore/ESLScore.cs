using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System
{
    public class ESLScore
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
        /// 參考修課紀錄ID(依據ESL2019寒假優化，成績參考修課紀錄ID，將廢除RefCourseID、RefStudentID)
        /// </summary>
        public string RefScAttendID { get; set; }

        /// <summary>
        /// 參考教師ID
        /// </summary>
        public string RefTeacherID { get; set; }

        /// <summary>
        /// 參考課程名稱
        /// </summary>
        public string RefCourseName { get; set; }

        /// <summary>
        /// 參考學生名稱
        /// </summary>
        public string RefStudentName { get; set; }

        /// <summary>
        /// 參考教師名稱
        /// </summary>
        public string RefTeacherName { get; set; }


        /// <summary>
        /// Term(評量)
        /// </summary>
        public string Term { get; set; }

        /// <summary>
        /// Subject(科目)
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Assessment(子項目)
        /// </summary>
        public string Assessment { get; set; }

        /// <summary>
        /// Custom_Assessment(教師自定義子項目)
        /// </summary>
        public string Custom_Assessment { get; set; }

        /// <summary>
        /// 舊成績值 (可能是分數:100、指標:G、評語:Good Job!， 本欄位作為紀錄LOG 前一筆成績使用)
        /// </summary>
        public string OldValue { get; set; }

    
        /// <summary>
        /// 成績值 (可能是分數:100、指標:G、評語:Good Job!)
        /// </summary>
        public string Value { get; set; }


        /// <summary>
        /// 成績分數 (經由系統按照設定比例計算後的分數)
        /// </summary>
        public decimal Score { get; set; }


        /// <summary>
        /// 是否有成績值 (評分教師是否有在Web輸入該成績)
        /// </summary>
        public bool HasValue { get; set; }


        /// <summary>
        /// 最後更新時間(比較重覆結構成績，以最後上傳為主)
        /// </summary>
        public string LastUpdate { get; set; }


        /// <summary>
        /// 指定Assessment成績分數比例(給特殊學生(ex:轉學生)成績使用，可以無視樣板的比例，強制將此成績以設定的Ratio　比例計算)
        /// </summary>
        public int? Ratio { get; set; }


    }
}
