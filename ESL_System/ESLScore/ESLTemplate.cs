using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System
{
    public class ESLTemplate
    {

        /// <summary>
        /// ESLTemplateID
        /// </summary>
        public string ID { get; set; }

        /// <summary>
        /// ESLTemplate名稱
        /// </summary>
        public string ESLTemplateName { get; set; }

        /// <summary>
        /// ESLTemplate子評量項目設定
        /// </summary>
        public string Description { get; set; }


        /// <summary>
        /// 轉換過的樣板設定階層        
        /// </summary>
        public List<Term> TermList = new List<Term>();

        

    }
}
