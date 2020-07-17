using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System.Model
{
    /// <summary>
    /// 學期成績資訊(加總) 計算用平均用 Model   (特定學期)
    /// </summary>
    class SemsTotalScoreInfo
    {
        public SemsTotalScoreInfo()
        {
            this.ListSubjects = new List<string>();
            this.TotalSubjScore = 0;
            this.TotalSubjGAP = 0;
        }

        public List<string> ListSubjects { get; }

        /// <summary>
        /// 累計科目數(Score)
        /// </summary>
        public int SubjectCountScore { get; set; }

        /// <summary>
        /// 累計科目數(GPA用)
        /// </summary>
        public int SubjectCountGPA { get; set; }

        /// <summary>
        /// 累計成績
        /// </summary>
        public decimal? TotalSubjScore { get; set; }

        /// <summary>
        /// 累計GPA
        /// </summary>
        public decimal? TotalSubjGAP { get; set; }



        /// <summary>
        /// 取得成績
        /// </summary>
        /// <param name="round">四捨五入小數點位數</param>
        /// <returns></returns>
        public decimal? GetScoreAvg(int round)
        {
            decimal? result = null;
            // 如果有科目
            if (this.SubjectCountScore != 0)
            {
                if (this.TotalSubjScore.HasValue) //  TotalScore 有值
                {
                    return Decimal.Round(this.TotalSubjScore.Value / this.SubjectCountScore, round, MidpointRounding.AwayFromZero);
                }
            }
            return result;
        }


        /// <summary>
        ///  取得學期成績算數平均(不處理四捨五入)
        /// </summary>
        /// <returns> </returns>
        public decimal? GetScoreAvg()
        {
            decimal? result = null;
            // 如果有科目
            if (this.SubjectCountScore != 0)
            {
                if (this.TotalSubjScore.HasValue) //  TotalScore 有值
                {
                    return this.TotalSubjScore.Value / this.SubjectCountScore;
                }
            }
            return result;
        }


        /// <summary>
        /// 取得學期GPA算數平均
        /// </summary>
        /// <param name="round">四捨五入小數點位數</param>
        /// <returns></returns>
        public decimal? GetGPAAvg(int round)
        {
            decimal? result = null;
            // 如果有科目
            if (this.SubjectCountGPA != 0)
            {
                if (this.TotalSubjGAP.HasValue) //  TotalScore 有值
                {
                    return Decimal.Round(this.TotalSubjGAP.Value / this.SubjectCountGPA, round, MidpointRounding.AwayFromZero);
                }
            }
            return result;
        }


        /// <summary>
        /// 取得學期GPA算術平均 (不處理算術平均)
        /// </summary>
        /// <returns> </returns>
        public decimal? GetGPAAvg()
        {
            decimal? result = null;
            // 如果有科目
            if (this.SubjectCountGPA != 0)
            {
                if (this.TotalSubjGAP.HasValue) //  TotalScore 有值
                {
                    return this.TotalSubjGAP.Value / this.SubjectCountGPA;
                }
            }
            return result;
        }


        /// <summary>
        /// 將成績++ (算術平均運算用)
        /// </summary>
        /// <param name="semsSubjScoreInfo"></param>
        public void AddScore(SemsSubjScoreInfo semsSubjScoreInfo)
        {
            this.ListSubjects.Add(semsSubjScoreInfo.Subject);

            if (semsSubjScoreInfo.SemsScore != null)
            {
                this.SubjectCountScore++;
                this.TotalSubjScore += semsSubjScoreInfo.SemsScore;
            }
            if (semsSubjScoreInfo.SemsGPA != null)
            {
                this.SubjectCountGPA++;
                this.TotalSubjGAP += semsSubjScoreInfo.SemsGPA;
            }
        }
    }
}
