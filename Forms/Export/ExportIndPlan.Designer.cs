namespace CafedralDB.Forms.Export
{
    partial class ExportIndPlan
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
			this.buttonExport = new System.Windows.Forms.Button();
			this.comboBoxYear = new System.Windows.Forms.ComboBox();
			this.studyYearBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.mainDBDataSet = new CafedralDB.MainDBDataSet();
			this.comboBoxTeacher = new System.Windows.Forms.ComboBox();
			this.employeeBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.studyYearTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.StudyYearTableAdapter();
			this.employeeTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.EmployeeTableAdapter();
			((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.employeeBindingSource)).BeginInit();
			this.SuspendLayout();
			// 
			// buttonExport
			// 
			this.buttonExport.Location = new System.Drawing.Point(100, 50);
			this.buttonExport.Name = "buttonExport";
			this.buttonExport.Size = new System.Drawing.Size(75, 23);
			this.buttonExport.TabIndex = 0;
			this.buttonExport.Text = "Экспорт";
			this.buttonExport.UseVisualStyleBackColor = true;
			this.buttonExport.Click += new System.EventHandler(this.buttonExport_Click);
			// 
			// comboBoxYear
			// 
			this.comboBoxYear.DataSource = this.studyYearBindingSource;
			this.comboBoxYear.DisplayMember = "StudyYear";
			this.comboBoxYear.FormattingEnabled = true;
			this.comboBoxYear.Location = new System.Drawing.Point(12, 26);
			this.comboBoxYear.Name = "comboBoxYear";
			this.comboBoxYear.Size = new System.Drawing.Size(121, 21);
			this.comboBoxYear.TabIndex = 1;
			this.comboBoxYear.ValueMember = "StudyYear";
			// 
			// studyYearBindingSource
			// 
			this.studyYearBindingSource.DataMember = "StudyYear";
			this.studyYearBindingSource.DataSource = this.mainDBDataSet;
			// 
			// mainDBDataSet
			// 
			this.mainDBDataSet.DataSetName = "MainDBDataSet";
			this.mainDBDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
			// 
			// comboBoxTeacher
			// 
			this.comboBoxTeacher.DataSource = this.employeeBindingSource;
			this.comboBoxTeacher.DisplayMember = "EmployeeName";
			this.comboBoxTeacher.FormattingEnabled = true;
			this.comboBoxTeacher.Location = new System.Drawing.Point(151, 26);
			this.comboBoxTeacher.Name = "comboBoxTeacher";
			this.comboBoxTeacher.Size = new System.Drawing.Size(121, 21);
			this.comboBoxTeacher.TabIndex = 2;
			this.comboBoxTeacher.ValueMember = "ID";
			// 
			// employeeBindingSource
			// 
			this.employeeBindingSource.DataMember = "Employee";
			this.employeeBindingSource.DataSource = this.mainDBDataSet;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(168, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(86, 13);
			this.label1.TabIndex = 3;
			this.label1.Text = "Преподаватель";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(39, 9);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(72, 13);
			this.label2.TabIndex = 4;
			this.label2.Text = "Учебный год";
			// 
			// studyYearTableAdapter
			// 
			this.studyYearTableAdapter.ClearBeforeFill = true;
			// 
			// employeeTableAdapter
			// 
			this.employeeTableAdapter.ClearBeforeFill = true;
			// 
			// ExportIndPlan
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(284, 81);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.comboBoxTeacher);
			this.Controls.Add(this.comboBoxYear);
			this.Controls.Add(this.buttonExport);
			this.Name = "ExportIndPlan";
			this.Text = "Экспорт индивидуального плана";
			this.Load += new System.EventHandler(this.ExportIndPlan_Load);
			((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.employeeBindingSource)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonExport;
        private System.Windows.Forms.ComboBox comboBoxYear;
        private System.Windows.Forms.ComboBox comboBoxTeacher;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private MainDBDataSet mainDBDataSet;
        private System.Windows.Forms.BindingSource studyYearBindingSource;
        private MainDBDataSetTableAdapters.StudyYearTableAdapter studyYearTableAdapter;
        private System.Windows.Forms.BindingSource employeeBindingSource;
        private MainDBDataSetTableAdapters.EmployeeTableAdapter employeeTableAdapter;
    }
}