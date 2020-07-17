using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System.Model
{ 
    /// <summary>
    /// 科目GPA對照表
    /// </summary>
    class SubjectGPAMapping
    {
        /// <summary>
        /// 科目名稱
        /// </summary>
        public string  SubjectName { get; set; }

        public bool IsHoner { get; set; }

        public bool IsAP { get; set; }

        public bool IsStandard { get; set; }

    }
}
