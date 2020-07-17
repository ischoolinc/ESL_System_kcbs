using FISCA.UDT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System.UDT
{
    [FISCA.UDT.TableName("ischool.mapping.score_gpa")]
    class ScoreGPAMapping : ActiveRecord
    {
        /// <summary>
        /// 成績區間最小值
        /// </summary>
        [Field(Field = "min_score", Indexed = false)]
        public decimal MinScore { get; set; }

        /// <summary>
        /// 成績區間最大值
        /// </summary>
        [Field(Field = "max_score", Indexed = false)]
        public decimal MaxScore { get; set; }


        /// <summary>
        /// 等第
        /// </summary>
        [Field(Field = "degree", Indexed = false)]
        public string Degree { get; set; }


        /// <summary>
        /// GPA
        /// </summary>
        [Field(Field = "gpa", Indexed = false)]
        public decimal? GPA { get; set; }

        /// <summary>
        /// Honers
        /// </summary>
        [Field(Field = "honers", Indexed = false)]
        public decimal? Honers { get; set; }


        /// <summary>
        /// Honers
        /// </summary>
        [Field(Field = "ap", Indexed = false)]
        public decimal? AP { get; set; }

        /// <summary>
        /// 組距
        /// </summary>
       
        public string RangString { get; set; }
    }
}
