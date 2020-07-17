using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System.Model
{
    /// <summary>
    /// 課程資訊 為了依照領域排序用
    /// </summary>
    class CourseInfo
    {
        public CourseInfo(string subject ,string domain,int order) {

            this.DomainOrder = 99999;
            this.Subject = subject;
            this.Domain = domain;
            this.DomainOrder = order;
        }
        public string Subject  { get; set; }
        public string Domain { get; set; }
        public int DomainOrder  { get; set; } 

    }
}
