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
    // 列印 WeeklyReport 的物件
    class WeeklyReportRecord
    {

        //
        // Summary:
        //     課程名稱
        public string CourseName { get; set; }

        //
        // Summary:
        //     科目
        public string Subject { get; set; }


        //
        // Summary:
        //     教師名稱
        public string TeacherName { get; set; }


        //
        // Summary:
        //     Score資料
        public string ScoreData { get; set; }

        //
        // Summary:
        //     Performance資料
        public string PerformanceData { get; set; }

        //
        // Summary:
        //     GeneralComment
        public string GeneralComment { get; set; }

        //
        // Summary:
        //     PersonalComment
        public string PersonalComment { get; set; }


    }
}
