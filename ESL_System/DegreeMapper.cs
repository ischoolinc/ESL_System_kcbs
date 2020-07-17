using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ESL_System
{
    public class DegreeMapper
    {
        // 處理等第轉換
        private Dictionary<decimal, string> _decimalToString = new Dictionary<decimal, string>();
        private Dictionary<string, decimal> _stringToDecimal = new Dictionary<string, decimal>();
        private List<decimal> _scoreList = new List<decimal>();

        public DegreeMapper()
        {
            #region 等第對照表
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["ESL等第對照表"];

            if (string.IsNullOrEmpty(cd["xml"]))
            {
                cd["xml"] = K12.Data.XmlHelper.LoadXml(Properties.Resources.預設ESL等第對照表).OuterXml;
                cd.Save();
            }

            if (cd.Contains("xml") && !string.IsNullOrEmpty(cd["xml"]))
            {
                XmlElement element = K12.Data.XmlHelper.LoadXml(cd["xml"]);

                foreach (XmlElement each in element.SelectNodes("ScoreMapping"))
                {
                    string degree = each.GetAttribute("EngName");
                    decimal score;
                    if (!decimal.TryParse(each.GetAttribute("Score"), out score))
                        score = 0;

                    if (!_decimalToString.ContainsKey(score))
                        _decimalToString.Add(score, degree);

                    if (!_stringToDecimal.ContainsKey(degree))
                        _stringToDecimal.Add(degree, score);
                }

                _scoreList = new List<decimal>();
                _scoreList.AddRange(_decimalToString.Keys);
                _scoreList.Sort(delegate (decimal a, decimal b)
                {
                    return b.CompareTo(a);
                });
            }
            #endregion

        }

        /// <summary>
        /// 取得等第
        /// </summary>
        /// <param name="score">分數</param>
        /// <returns>等第</returns>
        public string GetDegreeByScore(decimal score)
        {
            foreach (decimal s in _scoreList)
                if (score >= s) return _decimalToString[s];

            return _decimalToString[_scoreList[_scoreList.Count - 1]];
        }



    }
}
