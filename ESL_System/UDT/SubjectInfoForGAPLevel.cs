using FISCA.UDT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System.UDT
{
    /// <summary>
    /// 
    /// </summary>
    [FISCA.UDT.TableName("ischool.junior.subject.GPA_level")]
    class SubjectInfoForGAPLevel : ActiveRecord
    {
        /// <summary>
        /// 科目名稱
        /// </summary>
        [Field(Field = "subject_name", Indexed = false)]
        public string Subject  { get; set; }


        /// <summary>
        /// 是否為 standard
        /// </summary>
        [Field(Field = "is_standard", Indexed = false)]
        public Boolean IsStandard { get; set; }

        /// <summary>
        /// 使否為 honer
        /// </summary>
        [Field(Field = "is_honer", Indexed = false)]
        public Boolean IsHoner  {get; set; }

        /// <summary>
        /// 使否為 ap
        /// </summary>
        [Field(Field = "is_ap", Indexed = false)]
        public Boolean IsAP { get; set; }

    }
}
