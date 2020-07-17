using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System.Model
{
    class SemsTotalScoreInfos
    {
        // 儲放學生學期科目算術平均資訊(即時運算用) 其結構為 <studentID,<subjectName,SemsScoreInfo>>
        private Dictionary<string, SemsTotalScoreInfo> _DicSemsScoreInfo;


        public SemsTotalScoreInfos()
        {
            _DicSemsScoreInfo = new Dictionary<string, SemsTotalScoreInfo>();

        }

        /// <summary>
        /// 
        /// </summary>
        public void AddforCoumpte(string studentID , SemsSubjScoreInfo semsSubjScoreInfo) 
        {
            if (!_DicSemsScoreInfo.ContainsKey(studentID)) 
            {
                _DicSemsScoreInfo.Add(studentID, new SemsTotalScoreInfo());
            }

            
            _DicSemsScoreInfo[studentID].AddScore(semsSubjScoreInfo);

        }


        public bool ContainStuID(string studentID)
        {
            return this._DicSemsScoreInfo.ContainsKey(studentID);

        }

        /// <summary>
        /// 取得所有 
        /// </summary>
        public Dictionary<string, SemsTotalScoreInfo> GetAllSemsTotalScoreInfo()
        {
            return this._DicSemsScoreInfo;
        }
    }
}
