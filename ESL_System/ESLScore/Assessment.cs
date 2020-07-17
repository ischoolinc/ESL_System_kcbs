using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System
{
    public class Assessment
    {
        /// <summary>
        /// 子評量名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 子評量分數比例
        /// </summary>
        public string Weight { get; set; }

        /// <summary>
        /// 子評量評分教師(教師一、教師二、教師三)
        /// </summary>
        public string TeacherSequence { get; set; }

        /// <summary>
        /// 子評量評分種類
        /// </summary>
        public string Type { get; set; }

        /// <summary>
        /// 子評量是否准許老師自訂小考
        /// </summary>
        public string AllowCustomAssessment { get; set; }

        /// <summary>
        /// 子評量在回中文成績系統計算評量成績分數種類(定期、平時)
        /// </summary>
        public string ExamScoreType { get; set; }

        /// <summary>
        /// 專門給 子評量名稱Comments 的欄位，用來規範教師可以在web 上的評語字數，預設200
        /// </summary>
        public string InputLimit { get; set; }

        /// <summary>
        /// 指標型代碼List
        /// </summary>
        public List<Indicators> IndicatorsList { get; set; }


    }
}
