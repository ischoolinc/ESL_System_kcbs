using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using K12.Data;


namespace ESL_System.Form
{
    public partial class ESLTransferStudentSelectForm : FISCA.Presentation.Controls.BaseForm
    {
        // 目標課程ID
        private string _targetCourseID;

        // 目標課程名稱
        private string _targetCourseName;

        // 目標修課紀錄ID
        private string _targetScAttendID;

        // 指標性評語 可以使用項目
        private List<Indicators> _targetIndicatorList = new List<Indicators>();
                      
        //  學生修課資料
        private List<K12.Data.SCAttendRecord> _scaList = new List<SCAttendRecord>();
        
        public ESLTransferStudentSelectForm(List<string> targetCourseIDs)
        {
            InitializeComponent();
                                            
            #region 取得學生修課資料

            _scaList = K12.Data.SCAttend.SelectByCourseIDs(targetCourseIDs);

            #endregion

            List<K12.Data.CourseRecord> cr = K12.Data.Course.SelectByIDs(targetCourseIDs);

            _targetCourseName = cr[0].Name;
            _targetCourseID = cr[0].ID;
                                   
            labelX1.Text = _targetCourseName +"請選擇欲輸入ESL成績的轉學生。";
           
            // 填入修課學生
            FillStudent();
        }
        

        private void FillStudent()
        {
            dataGridViewX1.Rows.Clear();
           
            List<ESLScore> eslScoreList = new List<ESLScore>();

            foreach (K12.Data.SCAttendRecord scar in _scaList)
            {
                // 若學生有修課紀錄， 但是目前 狀態 為非一般，則不顯示。
                if (scar.Student.Status != StudentRecord.StudentStatus.一般)
                {
                    continue;
                }

                DataGridViewRow row = new DataGridViewRow();

                row.CreateCells(dataGridViewX1);
               
                row.Cells[0].Value = scar.Student.Class != null ? scar.Student.Class.Name : ""; // 學生班級
                row.Cells[1].Value = scar.Student != null ? "" + scar.Student.SeatNo : "";  // 學生座號
                row.Cells[2].Value = scar.Student != null ? "" + scar.Student.Name : "";      // 學生姓名
                row.Cells[3].Value = scar.Student != null ? "" + scar.Student.StudentNumber : "";  // 學生學號
                
                row.Tag = scar.ID;  // row tag 用sc_attend_id 就夠(依據2019 ESL 寒假優化項目)

                dataGridViewX1.Rows.Add(row);
            }

            // 依   學號 排序 (同Web 的成績輸入介面)
            dataGridViewX1.Sort(ColStudentNumber, ListSortDirection.Ascending);
        }

        // 離開
        private void buttonX1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridViewX1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            if (e.ColumnIndex < 0) return;
            if (e.RowIndex < 0) return;
            DataGridViewCell cell = dataGridViewX1.Rows[e.RowIndex].Cells[e.ColumnIndex];

            _targetScAttendID = "" + cell.OwningRow.Tag; //  targetTermName

            Form.ESLTransferScoreInputForm inputForm = new ESLTransferScoreInputForm(_targetScAttendID);

            inputForm.ShowDialog();

        }
    }
}
