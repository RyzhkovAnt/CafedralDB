namespace CafedralDB.Forms.Export
{
    partial class ExportContract
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.button1 = new System.Windows.Forms.Button();
            this.Teacher_comboBox = new System.Windows.Forms.ComboBox();
            this.EmployeeBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.mainDBDataSet1 = new CafedralDB.MainDBDataSet();
            this.Year_comboBox = new System.Windows.Forms.ComboBox();
            this.StudyYearBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.Semester_comboBox = new System.Windows.Forms.ComboBox();
            this.employeeTableAdapter1 = new CafedralDB.MainDBDataSetTableAdapters.EmployeeTableAdapter();
            this.studyYearTableAdapter1 = new CafedralDB.MainDBDataSetTableAdapters.StudyYearTableAdapter();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.EmployeeBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.StudyYearBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(128, 97);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(119, 38);
            this.button1.TabIndex = 0;
            this.button1.Text = "Создать";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Teacher_comboBox
            // 
            this.Teacher_comboBox.DataSource = this.EmployeeBindingSource;
            this.Teacher_comboBox.DisplayMember = "EmployeeName";
            this.Teacher_comboBox.FormattingEnabled = true;
            this.Teacher_comboBox.Location = new System.Drawing.Point(31, 25);
            this.Teacher_comboBox.Name = "Teacher_comboBox";
            this.Teacher_comboBox.Size = new System.Drawing.Size(121, 21);
            this.Teacher_comboBox.TabIndex = 1;
            this.Teacher_comboBox.ValueMember = "ID";
            // 
            // EmployeeBindingSource
            // 
            this.EmployeeBindingSource.DataMember = "Employee";
            this.EmployeeBindingSource.DataSource = this.mainDBDataSet1;
            // 
            // mainDBDataSet1
            // 
            this.mainDBDataSet1.DataSetName = "MainDBDataSet";
            this.mainDBDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // Year_comboBox
            // 
            this.Year_comboBox.DataSource = this.StudyYearBindingSource;
            this.Year_comboBox.DisplayMember = "StudyYear";
            this.Year_comboBox.FormattingEnabled = true;
            this.Year_comboBox.Location = new System.Drawing.Point(225, 25);
            this.Year_comboBox.Name = "Year_comboBox";
            this.Year_comboBox.Size = new System.Drawing.Size(121, 21);
            this.Year_comboBox.TabIndex = 2;
            this.Year_comboBox.ValueMember = "StudyYear";
            // 
            // StudyYearBindingSource
            // 
            this.StudyYearBindingSource.DataMember = "StudyYear";
            this.StudyYearBindingSource.DataSource = this.mainDBDataSet1;
            // 
            // Semester_comboBox
            // 
            this.Semester_comboBox.FormattingEnabled = true;
            this.Semester_comboBox.Items.AddRange(new object[] {
            "Осенний",
            "Весенний"});
            this.Semester_comboBox.Location = new System.Drawing.Point(225, 70);
            this.Semester_comboBox.Name = "Semester_comboBox";
            this.Semester_comboBox.Size = new System.Drawing.Size(121, 21);
            this.Semester_comboBox.TabIndex = 3;
            this.Semester_comboBox.Visible = false;
            // 
            // employeeTableAdapter1
            // 
            this.employeeTableAdapter1.ClearBeforeFill = true;
            // 
            // studyYearTableAdapter1
            // 
            this.studyYearTableAdapter1.ClearBeforeFill = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(48, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Преподаватель";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(254, 6);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(72, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Учебный год";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // ExportContract
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(376, 151);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Semester_comboBox);
            this.Controls.Add(this.Year_comboBox);
            this.Controls.Add(this.Teacher_comboBox);
            this.Controls.Add(this.button1);
            this.Name = "ExportContract";
            this.Text = "ExportContract";
            this.Load += new System.EventHandler(this.ExportContract_Load);
            ((System.ComponentModel.ISupportInitialize)(this.EmployeeBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.StudyYearBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox Teacher_comboBox;
        private System.Windows.Forms.ComboBox Year_comboBox;
        private System.Windows.Forms.ComboBox Semester_comboBox;
        private MainDBDataSet mainDBDataSet1;
        private MainDBDataSetTableAdapters.EmployeeTableAdapter employeeTableAdapter1;
        private MainDBDataSetTableAdapters.StudyYearTableAdapter studyYearTableAdapter1;
        private System.Windows.Forms.BindingSource EmployeeBindingSource;
        private System.Windows.Forms.BindingSource StudyYearBindingSource;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}