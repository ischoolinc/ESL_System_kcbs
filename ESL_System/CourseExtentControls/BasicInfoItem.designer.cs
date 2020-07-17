using Framework;
using System;
namespace ESL_System.CourseExtendControls
{
    partial class BasicInfoItem
    {
        /// <summary> 
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該公開 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
                //Teacher.Instance.TeacherDataChanged -= new EventHandler<TeacherDataChangedEventArgs>(Instance_TeacherDataChanged);
                //Teacher.Instance.TeacherInserted -= new EventHandler(Instance_TeacherInserted);
                //Teacher.Instance.TeacherDeleted -= new EventHandler<TeacherDeletedEventArgs>(Instance_TeacherDeleted);
                //ClassEntity.Instance.ClassInserted-= new EventHandler<InsertClassEventArgs>(Instance_ClassInserted);
                //ClassEntity.Instance.ClassUpdated -= new EventHandler<UpdateClassEventArgs>(Instance_ClassUpdated);
                //ClassEntity.Instance.ClassDeleted -= new EventHandler<DeleteClassEventArgs>(Instance_ClassDeleted);
                //CourseEntity.Instance.ForeignTableChanged -= new EventHandler(Instance_ForeignTableChanged);
                //CourseEntity.Instance.CourseChanged -= new EventHandler<CourseChangeEventArgs>(Instance_CourseChanged);
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。
        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.lbDifficulty = new DevComponents.DotNetBar.LabelX();
            this.txtDifficulty = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.lbClassroom = new DevComponents.DotNetBar.LabelX();
            this.txtClassroom = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.SuspendLayout();
            // 
            // lbDifficulty
            // 
            // 
            // 
            // 
            this.lbDifficulty.BackgroundStyle.Class = "";
            this.lbDifficulty.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbDifficulty.Location = new System.Drawing.Point(3, 2);
            this.lbDifficulty.Name = "lbDifficulty";
            this.lbDifficulty.Size = new System.Drawing.Size(104, 23);
            this.lbDifficulty.TabIndex = 2;
            this.lbDifficulty.Text = "課程難度(Level)";
            this.lbDifficulty.TextAlignment = System.Drawing.StringAlignment.Far;
            // 
            // txtDifficulty
            // 
            // 
            // 
            // 
            this.txtDifficulty.Border.Class = "TextBoxBorder";
            this.txtDifficulty.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtDifficulty.Location = new System.Drawing.Point(113, 2);
            this.txtDifficulty.MaxLength = 50;
            this.txtDifficulty.Name = "txtDifficulty";
            this.txtDifficulty.Size = new System.Drawing.Size(151, 25);
            this.txtDifficulty.TabIndex = 3;
            this.txtDifficulty.TextChanged += new System.EventHandler(this.txtDifficulty_TextChanged);
            // 
            // lbClassroom
            // 
            // 
            // 
            // 
            this.lbClassroom.BackgroundStyle.Class = "";
            this.lbClassroom.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbClassroom.Location = new System.Drawing.Point(288, 2);
            this.lbClassroom.Name = "lbClassroom";
            this.lbClassroom.Size = new System.Drawing.Size(70, 23);
            this.lbClassroom.TabIndex = 4;
            this.lbClassroom.Text = "上課地點";
            this.lbClassroom.TextAlignment = System.Drawing.StringAlignment.Far;
            // 
            // txtClassroom
            // 
            // 
            // 
            // 
            this.txtClassroom.Border.Class = "TextBoxBorder";
            this.txtClassroom.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtClassroom.Location = new System.Drawing.Point(364, 2);
            this.txtClassroom.MaxLength = 50;
            this.txtClassroom.Name = "txtClassroom";
            this.txtClassroom.Size = new System.Drawing.Size(151, 25);
            this.txtClassroom.TabIndex = 5;
            this.txtClassroom.TextChanged += new System.EventHandler(this.txtClassroom_TextChanged);
            // 
            // BasicInfoItem
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.lbClassroom);
            this.Controls.Add(this.txtClassroom);
            this.Controls.Add(this.lbDifficulty);
            this.Controls.Add(this.txtDifficulty);
            this.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MinimumSize = new System.Drawing.Size(550, 0);
            this.Name = "BasicInfoItem";
            this.Size = new System.Drawing.Size(550, 40);
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.LabelX lbDifficulty;
        protected DevComponents.DotNetBar.Controls.TextBoxX txtDifficulty;
        private DevComponents.DotNetBar.LabelX lbClassroom;
        protected DevComponents.DotNetBar.Controls.TextBoxX txtClassroom;
    }
}
