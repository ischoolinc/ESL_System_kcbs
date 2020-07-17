using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ESL_System
{
    // 2018/6/8 穎驊　新增此UDT 用來記錄每一個ESL 評分樣版　其每學年度、學期　不同階段的(期中、期末、學期)　報表的樣版
    [FISCA.UDT.TableName("esl.report.template")]
    public class UDT_ReportTemplate : FISCA.UDT.ActiveRecord
    {
        public UDT_ReportTemplate()
        {

        }

        /// <summary>
        /// 參考的評分樣版 ID
        /// </summary>
        [FISCA.UDT.Field]
        public string Ref_exam_Template_ID { get; set; }

        /// <summary>
        /// 設定檔名稱
        /// </summary>
        [FISCA.UDT.Field]
        public string Name { get; set; }

        /// <summary>
        /// 學年度
        /// </summary>
        [FISCA.UDT.Field]
        public string SchoolYear { get; set; }

        /// <summary>
        /// 學期
        /// </summary>
        [FISCA.UDT.Field]
        public string Semester { get; set; }

        /// <summary>
        /// 試別
        /// </summary>
        [FISCA.UDT.Field]
        public string Exam { get; set; }

        /// <summary>
        /// 列印樣板
        /// </summary>
        [FISCA.UDT.Field]
        private string TemplateStream { get; set; }

        public Aspose.Words.Document Template { get; set; }
      
        /// <summary>
        /// 在儲存前，把資料填入儲存欄位中
        /// </summary>
        public void Encode()
        {           
            System.IO.MemoryStream stream = new System.IO.MemoryStream();
            this.Template.Save(stream, Aspose.Words.SaveFormat.Doc);
            this.TemplateStream = Convert.ToBase64String(stream.ToArray());
        }
        /// <summary>
        /// 在資料取出後，把資料從儲存欄位轉換至資料欄位
        /// </summary>
        public void Decode()
        {            
            this.Template = new Aspose.Words.Document(new MemoryStream(Convert.FromBase64String(this.TemplateStream)));
        }
    }
}
