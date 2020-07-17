using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System
{
    class Global
    {

        /// <summary>
        /// 取得獎懲名稱
        /// </summary>
        /// <returns></returns>
        public static List<string> GetDisciplineNameList()
        {
            return new string[] { "大功", "小功", "嘉獎", "大過", "小過", "警告" }.ToList();
        }

    }
}
