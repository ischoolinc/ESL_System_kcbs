using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA;
using FISCA.Presentation;
using K12.Presentation;
using FISCA.Permission;
using JHSchool;



namespace ESL_System
{
    public class Program
    {
        //2018/4/11 穎驊因應康橋英文系統、弘文ESL 專案 ，開始建構教務作業ESL 評分樣版設定
        [FISCA.MainMethod()]
        public static void Main()
        {

            FISCA.UDT.AccessHelper accessHelper = new FISCA.UDT.AccessHelper();

            accessHelper.Select<UDT_ReportTemplate>(); // 先將UDT 選起來，如果是第一次開啟沒有話就會新增
            accessHelper.Select<UDT_WeeklyReportTemplate>(); // 先將UDT 選起來，如果是第一次開啟沒有話就會新增

            Catalog ribbon = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            ribbon.Add(new RibbonFeature("ESL評分樣版設定", "ESL評分樣版設定"));

            MotherForm.RibbonBarItems["教務作業", "基本設定"]["設定"]["ESL評分樣版設定"].Enable = UserAcl.Current["ESL評分樣版設定"].Executable;

            MotherForm.RibbonBarItems["教務作業", "基本設定"]["設定"]["ESL評分樣版設定"].Click += delegate
            {
                Form.ESL_TemplateSetupManager form = new Form.ESL_TemplateSetupManager();

                form.ShowDialog();

            };

            Catalog ribbon2 = RoleAclSource.Instance["課程"]["ESL課程"];
            ribbon2.Add(new RibbonFeature("ESL評量分數計算", "評量成績結算"));

            MotherForm.RibbonBarItems["課程", "ESL課程"]["評量成績結算"].Enable = false;

            K12.Presentation.NLDPanels.Course.SelectedSourceChanged += (sender, e) =>
            {
                if (K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0)
                {
                    MotherForm.RibbonBarItems["課程", "ESL課程"]["評量成績結算"].Enable = UserAcl.Current["ESL評量分數計算"].Executable;
                }
                else
                {
                    MotherForm.RibbonBarItems["課程", "ESL課程"]["評量成績結算"].Enable = false;
                }
            };

            MotherForm.RibbonBarItems["課程", "ESL課程"]["評量成績結算"].Image = Properties.Resources.calc_64;
            MotherForm.RibbonBarItems["課程", "ESL課程"]["評量成績結算"].Size = RibbonBarButton.MenuButtonSize.Medium;

            MotherForm.RibbonBarItems["課程", "ESL課程"]["評量成績結算"].Click += delegate
            {
                Form.CheckCalculateTermForm form = new Form.CheckCalculateTermForm(K12.Presentation.NLDPanels.Course.SelectedSource);

                form.ShowDialog();

            };


            //2019/02/26 穎驊註解， 依據ESL 寒假優化項目 課程上的 ESL 報表功能 將移除，日後會統一在學生上列印

            //Catalog ribbon3 = RoleAclSource.Instance["課程"]["ESL報表"];
            //ribbon3.Add(new RibbonFeature("ESL成績單", "ESL報表"));

            //MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["ESL成績單"].Enable = UserAcl.Current["ESL成績單"].Executable && K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0;

            //K12.Presentation.NLDPanels.Course.SelectedSourceChanged += delegate
            //{
            //    MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["ESL成績單"].Enable = UserAcl.Current["ESL成績單"].Executable && (K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0);
            //};


            //MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["ESL成績單"].Click += delegate
            //{
            //    List<string> eslCouseList = K12.Presentation.NLDPanels.Course.SelectedSource.ToList();

            //    Form.PrintESLReportForm printform = new Form.PrintESLReportForm(eslCouseList);

            //    printform.ShowDialog();

            //};


            Catalog ribbon4 = RoleAclSource.Instance["課程"]["ESL課程"];
            ribbon4.Add(new RibbonFeature("ESL課程成績輸入狀況", "成績輸入狀況"));

            MotherForm.RibbonBarItems["課程", "ESL課程"]["成績輸入狀況"].Enable = false;

            K12.Presentation.NLDPanels.Course.SelectedSourceChanged += (sender, e) =>
            {
                if (K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0)
                {
                    MotherForm.RibbonBarItems["課程", "ESL課程"]["成績輸入狀況"].Enable = UserAcl.Current["ESL課程成績輸入狀況"].Executable;
                }
                else
                {
                    MotherForm.RibbonBarItems["課程", "ESL課程"]["成績輸入狀況"].Enable = false;
                }
            };

            MotherForm.RibbonBarItems["課程", "ESL課程"]["成績輸入狀況"].Image = Properties.Resources.calc_64;
            MotherForm.RibbonBarItems["課程", "ESL課程"]["成績輸入狀況"].Size = RibbonBarButton.MenuButtonSize.Medium;

            MotherForm.RibbonBarItems["課程", "ESL課程"]["成績輸入狀況"].Click += delegate
            {

                List<string> eslCouseList = K12.Presentation.NLDPanels.Course.SelectedSource.ToList();

                Form.ESLCourseScoreStatusForm form = new Form.ESLCourseScoreStatusForm(eslCouseList);

                form.ShowDialog();

            };


            Catalog ribbon5 = RoleAclSource.Instance["學生"]["報表"];
            ribbon5.Add(new RibbonFeature("1C389099-FBA2-4C4B-9C0C-0FD7CB18EBC3", "ESL個人成績單"));

            MotherForm.RibbonBarItems["學生", "資料統計"]["報表"]["ESL報表"]["ESL個人成績單"].Enable = UserAcl.Current["1C389099-FBA2-4C4B-9C0C-0FD7CB18EBC3"].Executable && K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0;

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                MotherForm.RibbonBarItems["學生", "資料統計"]["報表"]["ESL報表"]["ESL個人成績單"].Enable = UserAcl.Current["1C389099-FBA2-4C4B-9C0C-0FD7CB18EBC3"].Executable && (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0);
            };


            MotherForm.RibbonBarItems["學生", "資料統計"]["報表"]["ESL報表"]["ESL個人成績單"].Click += delegate
            {
                Form.PrintStudentESLReportForm form = new Form.PrintStudentESLReportForm();

                form.ShowDialog();
            };


            Catalog ribbon6 = RoleAclSource.Instance["課程"]["ESL課程"];
            ribbon6.Add(new RibbonFeature("12813482-3D73-4AEF-8924-FA5189C9BDE3", "課程成績匯出"));

            MotherForm.RibbonBarItems["課程", "ESL課程"]["課程成績匯出"].Enable = false;

            K12.Presentation.NLDPanels.Course.SelectedSourceChanged += (sender, e) =>
            {
                if (K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0)
                {
                    MotherForm.RibbonBarItems["課程", "ESL課程"]["課程成績匯出"].Enable = UserAcl.Current["12813482-3D73-4AEF-8924-FA5189C9BDE3"].Executable;
                }
                else
                {
                    MotherForm.RibbonBarItems["課程", "ESL課程"]["課程成績匯出"].Enable = false;
                }
            };

            MotherForm.RibbonBarItems["課程", "ESL課程"]["課程成績匯出"].Image = Properties.Resources.admissions_zoom_64;
            MotherForm.RibbonBarItems["課程", "ESL課程"]["課程成績匯出"].Size = RibbonBarButton.MenuButtonSize.Medium;

            MotherForm.RibbonBarItems["課程", "ESL課程"]["課程成績匯出"].Click += delegate
            {

                List<string> eslCouseList = K12.Presentation.NLDPanels.Course.SelectedSource.ToList();

                // 舊課程成績匯出，其格式 為SQL 抓出來的資料Excel， 不符合使用者直觀的介面，
                // 將用新的介面功能取代(從取前N名功能改寫過來)
                ExportESLscore_New exporter = new ExportESLscore_New(eslCouseList);
                exporter.PrintReport();


            };

            //MotherForm.RibbonBarItems["課程", "ESL課程"]["匯入新竹成績(暫時)"].Click += delegate
            //{
            //    ImportHCScore import = new ImportHCScore();
            //};

            Catalog ribbon7 = RoleAclSource.Instance["課程"]["ESL報表"];
            ribbon7.Add(new RibbonFeature("EB1AF12A-F5B3-41B9-9F9A-B2ABE541F3A3", "班級前N名"));

            MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["班級前N名"].Enable = UserAcl.Current["EB1AF12A-F5B3-41B9-9F9A-B2ABE541F3A3"].Executable && K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0;

            K12.Presentation.NLDPanels.Course.SelectedSourceChanged += delegate
            {
                MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["班級前N名"].Enable = UserAcl.Current["EB1AF12A-F5B3-41B9-9F9A-B2ABE541F3A3"].Executable && (K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0);
            };

            MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["班級前N名"].Click += delegate
            {
                List<string> eslCouseList = K12.Presentation.NLDPanels.Course.SelectedSource.ToList();


                Form.PrintESLAwardTotalScoreReportForm printform = new Form.PrintESLAwardTotalScoreReportForm(eslCouseList);

                printform.ShowDialog();
            };


            Catalog ribbon8 = RoleAclSource.Instance["課程"]["ESL報表"];
            ribbon8.Add(new RibbonFeature("4B1318A1-1DA6-4A23-8508-2A394CFE4D9C", "進步名單"));

            MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["進步名單"].Enable = UserAcl.Current["4B1318A1-1DA6-4A23-8508-2A394CFE4D9C"].Executable && K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0;

            K12.Presentation.NLDPanels.Course.SelectedSourceChanged += delegate
            {
                MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["進步名單"].Enable = UserAcl.Current["4B1318A1-1DA6-4A23-8508-2A394CFE4D9C"].Executable && (K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0);
            };

            MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["進步名單"].Click += delegate
            {
                List<string> eslCouseList = K12.Presentation.NLDPanels.Course.SelectedSource.ToList();

                Form.PrintESLAwardProgressScoreReportForm printform = new Form.PrintESLAwardProgressScoreReportForm(eslCouseList);

                printform.ShowDialog();
            };


            Catalog ribbon9 = RoleAclSource.Instance["課程"]["ESL課程"];
            ribbon4.Add(new RibbonFeature("DA908268-F1E6-4514-B69D-386EC37184A9", "缺考成績處理(指定比例)"));

            MotherForm.RibbonBarItems["課程", "ESL課程"]["缺考成績處理(指定比例)"].Enable = false;

            K12.Presentation.NLDPanels.Course.SelectedSourceChanged += (sender, e) =>
            {
                // 一次只能選一個 課程 補輸入 轉班生成績
                if (K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0 & K12.Presentation.NLDPanels.Course.SelectedSource.Count < 2)
                {
                    MotherForm.RibbonBarItems["課程", "ESL課程"]["缺考成績處理(指定比例)"].Enable = UserAcl.Current["DA908268-F1E6-4514-B69D-386EC37184A9"].Executable;
                }
                else
                {
                    MotherForm.RibbonBarItems["課程", "ESL課程"]["缺考成績處理(指定比例)"].Enable = false;
                }
            };

            MotherForm.RibbonBarItems["課程", "ESL課程"]["缺考成績處理(指定比例)"].Image = Properties.Resources.calc_64;
            MotherForm.RibbonBarItems["課程", "ESL課程"]["缺考成績處理(指定比例)"].Size = RibbonBarButton.MenuButtonSize.Medium;

            MotherForm.RibbonBarItems["課程", "ESL課程"]["缺考成績處理(指定比例)"].Click += delegate
            {

                List<string> eslCouseIDList = K12.Presentation.NLDPanels.Course.SelectedSource;

                Form.ESLTransferStudentSelectForm form = new Form.ESLTransferStudentSelectForm(eslCouseIDList);

                form.ShowDialog();
            };


            //課程基本資訊 (課程難度(Level)、上課地點) 
            // 此外掛方式，會讓只有 加掛 ESL 模組學校  才可以在 課程上 有 課程難度(Level)、上課地點 兩個欄位可以編輯
            K12.Presentation.NLDPanels.Course.AddDetailBulider(new FISCA.Presentation.DetailBulider<CourseExtendControls.BasicInfoItem>());


            Catalog ribbon10 = RoleAclSource.Instance["學生"]["報表"];
            ribbon10.Add(new RibbonFeature("1C26A90B-DD2E-4298-A6D8-35C7E3C8EC7D", "WeeklyReport報表"));

            MotherForm.RibbonBarItems["學生", "資料統計"]["報表"]["ESL報表"]["WeeklyReport報表"].Enable = UserAcl.Current["1C26A90B-DD2E-4298-A6D8-35C7E3C8EC7D"].Executable && K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0;

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                MotherForm.RibbonBarItems["學生", "資料統計"]["報表"]["ESL報表"]["WeeklyReport報表"].Enable = UserAcl.Current["1C26A90B-DD2E-4298-A6D8-35C7E3C8EC7D"].Executable && (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0);
            };


            MotherForm.RibbonBarItems["學生", "資料統計"]["報表"]["ESL報表"]["WeeklyReport報表"].Click += delegate
            {
                Form.PrintWeeklyReportForm form = new Form.PrintWeeklyReportForm();

                form.ShowDialog();
            };


        }
    }
}
