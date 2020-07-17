
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.UDT;
using FISCA.Presentation.Controls;
using System.IO;
using FISCA.Data;
using K12.Data;
using FISCA.DSAUtil;
using FISCA.Authentication;
using FISCA.LogAgent;

namespace ESL_System.CourseExtendControls
{
    //[Framework.AccessControl.FeatureCode("Content0200")]    
    [FISCA.Permission.FeatureCode("JHSchool.Course.Detail0000", "基本資料")]
    internal partial class BasicInfoItem : FISCA.Presentation.DetailContent
    {
        DataValueManager _ValueManager;
        CourseBaseLogMachine machine = new CourseBaseLogMachine();
        private bool _initialing = false;

        public BasicInfoItem()
        {
            InitializeComponent();

            Group = "基本資料";
            _ValueManager = new DataValueManager();
        }

        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            SaveButtonVisible = false;
            CancelButtonVisible = false;

            _initialing = true;

            #region 取得difficulty及classroom資料的sql
            string selectSql = @"
SELECT
    difficulty
    , classroom
FROM
    course
WHERE
    id = " + PrimaryKey + @"
";
            #endregion

            QueryHelper queryHelper = new QueryHelper();
            DataTable dt = queryHelper.Select(selectSql);
            txtDifficulty.Text = "" + dt.Rows[0]["difficulty"];
            txtClassroom.Text = "" + dt.Rows[0]["classroom"];

            WatchValue("Difficulty", txtDifficulty.Text);
            WatchValue("Classroom", txtClassroom.Text);

            _initialing = false;

            machine.AddBefore(lbDifficulty.Text.Replace("　", "").Replace(" ", ""), txtDifficulty.Text);
            machine.AddBefore(lbClassroom.Text.Replace("　", "").Replace(" ", ""), txtClassroom.Text);
        }

        private void WatchValue(string name, object value)
        {
            if (value is String) //如果是「字串」就一般方法處理。
                _ValueManager.AddValue(name, value.ToString());
            else
            { //非字串用「物件」方式處理。
                if (value != null)
                    _ValueManager.AddValue(name, value.GetHashCode().ToString());
                else
                    _ValueManager.AddValue(name, "");
            }
        }

        protected override void OnSaveButtonClick(EventArgs e)
        {
            try
            {
                Dictionary<string, string> items = _ValueManager.GetDirtyItems();

                if (items.Count <= 0) //沒有任何更動。
                    return;

                machine.AddAfter(lbDifficulty.Text.Replace("　", "").Replace(" ", ""), txtDifficulty.Text);
                machine.AddAfter(lbClassroom.Text.Replace("　", "").Replace(" ", ""), txtClassroom.Text);

                bool _update_required = false;
                if (items.ContainsKey("Difficulty"))
                {
                    #region 更新difficulty欄位的sql
                    string updateDifficulty = @"
UPDATE
    course
SET
    difficulty = '" + txtDifficulty.Text.TrimStart(' ').TrimEnd(' ').Replace("'", "''") + @"'
WHERE
    id = " + PrimaryKey + @"
";
                    #endregion
                    UpdateHelper updateHelper = new UpdateHelper();
                    updateHelper.Execute(updateDifficulty);
                    _update_required = true;
                }

                if (items.ContainsKey("Classroom"))
                {
                    #region 更新classroom欄位的sql
                    string updateClassroom = @"
UPDATE
    course
SET
    classroom = '" + txtClassroom.Text.TrimStart(' ').TrimEnd(' ').Replace("'", "''") + @"'
WHERE
    id = " + PrimaryKey + @"
";
                    #endregion
                    UpdateHelper updateHelper = new UpdateHelper();
                    updateHelper.Execute(updateClassroom);
                    _update_required = true;
                }

                if (_update_required)
                {
                    #region Log

                    StringBuilder desc = new StringBuilder("");
                    desc.AppendLine("課程名稱：" + Course.SelectByID(PrimaryKey).Name + " ");
                    desc.AppendLine(machine.GetDescription());
                    string actor = DSAServices.UserAccount;
                    string client_info = ClientInfo.GetCurrentClientInfo().OutputResult().OuterXml;

                    #region 寫Log的SQL
                    string insertLogSql = @"
INSERT INTO log
(
	actor
	, action_type
	, action
	, target_id
	, server_time
	, client_info
	, action_by
	, description
)
VALUES
(
	'" + actor + @"'
	, 'Import'
	, '修改'
	, " + PrimaryKey + @"
	, now()
	, '" + client_info + @"'
	, '系統歷程'   
	, '" + desc + @"'
)
";
                    #endregion

                    UpdateHelper updateHelper = new UpdateHelper();
                    updateHelper.Execute(insertLogSql);

                    #endregion

                }

                SaveButtonVisible = false;
                CancelButtonVisible = false;

                WatchValue("Difficulty", txtDifficulty.Text);
                WatchValue("Classroom", txtClassroom.Text);
            }
            catch (Exception ex)
            {
                MsgBox.Show(ex.Message);
            }
        }

        protected override void OnCancelButtonClick(EventArgs e)
        {
            txtDifficulty.Text = _ValueManager.GetOldValue("Difficulty");
            txtClassroom.Text = _ValueManager.GetOldValue("Classroom");
        }

        private void txtDifficulty_TextChanged(object sender, EventArgs e)
        {
            if (!_initialing)
                OnValueChanged("Difficulty", txtDifficulty.Text);
        }

        private void txtClassroom_TextChanged(object sender, EventArgs e)
        {
            if (!_initialing)
                OnValueChanged("Classroom", txtClassroom.Text);
        }

        protected void OnValueChanged(string name, string value)
        {
            _ValueManager.SetValue(name, value);
            RaiseEvent();
        }

        protected void RaiseEvent()
        {
            if (_ValueManager.IsDirty != SaveButtonVisible)
            {
                SaveButtonVisible = _ValueManager.IsDirty;
                CancelButtonVisible = _ValueManager.IsDirty;
                if (this.SaveButtonVisibleChanged != null)
                {
                    SaveButtonVisibleChanged.Invoke(this, new EventArgs());
                }
                if (this.CancelButtonVisibleChanged != null)
                {
                    CancelButtonVisibleChanged.Invoke(this, new EventArgs());
                }
            }
        }

        public event EventHandler SaveButtonVisibleChanged;

        public event EventHandler CancelButtonVisibleChanged;

        class CourseBaseLogMachine
        {
            Dictionary<string, string> beforeData = new Dictionary<string, string>();
            Dictionary<string, string> afterData = new Dictionary<string, string>();

            public void AddBefore(string key, string value)
            {
                if (!beforeData.ContainsKey(key))
                    beforeData.Add(key, value);
                else
                    beforeData[key] = value;
            }

            public void AddAfter(string key, string value)
            {
                if (!afterData.ContainsKey(key))
                    afterData.Add(key, value);
                else
                    afterData[key] = value;
            }

            public string GetDescription()
            {
                //「」
                StringBuilder desc = new StringBuilder("");

                foreach (string key in beforeData.Keys)
                {
                    if (afterData.ContainsKey(key) && afterData[key] != beforeData[key])
                    {
                        desc.AppendLine("欄位「" + key + "」由「" + beforeData[key] + "」變更為「" + afterData[key] + "」");
                    }
                }

                return desc.ToString();
            }
        }

        public class DataValueManager
        {
            private Dictionary<string, string> _displayTexts;
            private Dictionary<string, string> _nowValues;
            private Dictionary<string, string> _oldValues;

            public DataValueManager()
            {
                Initialize();
            }

            public void AddValue(string name, string value)
            {
                AddValue(name, value, name);
            }

            /// <summary>
            /// 加入項目
            /// </summary>
            /// <param name="name">項目索引</param>
            /// <param name="value">項目值</param>
            public void AddValue(string name, string value, string displayText)
            {
                if (_nowValues.ContainsKey(name))
                    _nowValues[name] = value;
                else
                    _nowValues.Add(name, value);

                if (_oldValues.ContainsKey(name))
                    _oldValues[name] = value;
                else
                    _oldValues.Add(name, value);

                if (_displayTexts.ContainsKey(name))
                    _displayTexts[name] = displayText;
                else
                    _displayTexts.Add(name, displayText);
            }

            /// <summary>
            /// 變更項目
            /// </summary>
            /// <param name="name">項目索引</param>
            /// <param name="value">新值</param>
            public void SetValue(string name, string value)
            {
                if (_nowValues.ContainsKey(name))
                    _nowValues[name] = value;
            }

            /// <summary>
            /// 取出目前所有項目
            /// </summary>
            /// <returns></returns>
            public Dictionary<string, string> GetValues()
            {
                return _nowValues;
            }

            /// <summary>
            /// 取出指定名稱的原始資料。
            /// </summary>
            public string GetOldValue(string name)
            {
                return _oldValues[name];
            }

            public string GetDisplayText(string name)
            {
                return _displayTexts[name];
            }

            /// <summary>
            /// 將所有項目清空重設
            /// </summary>
            public void ResetValues()
            {
                Initialize();
            }

            /// <summary>
            /// 將變更項目設為預設項目
            /// </summary>
            public void MakeDirtyToClean()
            {
                foreach (string key in _nowValues.Keys)
                {
                    _oldValues[key] = _nowValues[key];
                }
            }

            /// <summary>
            /// 判斷是否已有值被變更
            /// </summary>
            public bool IsDirty
            {
                get
                {
                    foreach (string key in _oldValues.Keys)
                    {
                        if (IsDirtyItem(key))
                            return true;
                    }
                    return false;
                }
            }

            /// <summary>
            /// 初始化
            /// </summary>
            private void Initialize()
            {
                _nowValues = new Dictionary<string, string>();
                _oldValues = new Dictionary<string, string>();
                _displayTexts = new Dictionary<string, string>();
            }

            public DSRequest GetRequest(string rootName, string dataElementName, string fieldElementName, string conditionElementName, string conditionName, string id)
            {
                DSRequest dsreq = new DSRequest();
                DSXmlHelper helper = new DSXmlHelper(rootName);
                if (!string.IsNullOrEmpty(dataElementName))
                {
                    helper.AddElement(dataElementName);
                    helper.AddElement(dataElementName, fieldElementName);
                    helper.AddElement(dataElementName, conditionElementName);
                    fieldElementName = dataElementName + "/" + fieldElementName;
                    conditionElementName = dataElementName + "/" + conditionElementName;
                }
                else
                {
                    helper.AddElement(fieldElementName);
                    helper.AddElement(conditionElementName);
                }

                foreach (string key in _nowValues.Keys)
                {
                    if (_nowValues[key] != _oldValues[key])
                    {
                        helper.AddElement(fieldElementName, key, _nowValues[key]);
                    }
                }

                helper.AddElement(conditionElementName, conditionName, id);
                dsreq.SetContent(helper);
                //Console.WriteLine(helper.GetRawXml());
                return dsreq;
            }

            /// <summary>
            /// 取出變更項目清單
            /// </summary>
            /// <returns>變更項目清單，key值為索引,value為變更後的值</returns>
            public Dictionary<string, string> GetDirtyItems()
            {
                Dictionary<string, string> dic = new Dictionary<string, string>();

                foreach (string key in _oldValues.Keys)
                {
                    if (IsDirtyItem(key))
                        dic.Add(key, _nowValues[key]);
                }
                return dic;
            }

            /// <summary>
            /// 判斷key值是否已變更
            /// </summary>
            /// <param name="key">索引</param>
            /// <returns>若已變更則傳回 true，反之傳回 false</returns>
            public bool IsDirtyItem(string key)
            {
                return _oldValues[key] != _nowValues[key];
            }
        }
    }
}