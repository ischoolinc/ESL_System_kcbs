using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA;
using FISCA.Presentation;
using K12.Presentation;
using FISCA.Permission;


namespace ESL_System_Kcbs_Report
{
    public class Program
    {
        //2018/5/16 穎驊因應康橋英文系統ESL 專案 ，開始建構課程 提供列印期末成績單
        [FISCA.MainMethod()]
        public static void Main()
        {
            Catalog ribbon = RoleAclSource.Instance["課程"]["ESL報表"];
            ribbon.Add(new RibbonFeature("康橋ESL期末成績單", "康橋ESL期末成績單"));

            MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["ESL期末成績單"].Enable = UserAcl.Current["ESL期末成績單"].Executable && K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0;

            K12.Presentation.NLDPanels.Course.SelectedSourceChanged += delegate
            {
                MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["ESL期末成績單"].Enable = UserAcl.Current["ESL期末成績單"].Executable && (K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0);
            };


            MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["ESL期末成績單"].Click += delegate
            {

                List<K12.Data.CourseRecord> esl_couse_list = K12.Data.Course.SelectByIDs(K12.Presentation.NLDPanels.Course.SelectedSource);

                ESL_KcbsFinalReportForm form = new ESL_KcbsFinalReportForm(esl_couse_list);





            };


        }
    }
}
