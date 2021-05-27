namespace CafedralDB.Forms.Export
{
    partial class ExportSemester
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
			this.comboBoxSemester = new System.Windows.Forms.ComboBox();
			this.semesterBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.mainDBDataSet = new CafedralDB.MainDBDataSet();
			this.comboBoxYear = new System.Windows.Forms.ComboBox();
			this.studyYearBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.mainDBDataSetBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.studyYearTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.StudyYearTableAdapter();
			this.semesterTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.SemesterTableAdapter();
			((System.ComponentModel.ISupportInitialize)(this.semesterBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSetBindingSource)).BeginInit();
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
			// comboBoxSemester
			// 
			this.comboBoxSemester.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBoxSemester.FormattingEnabled = true;
			this.comboBoxSemester.Items.AddRange(new object[] {
            "Осенний",
            "Весенний"});
			this.comboBoxSemester.Location = new System.Drawing.Point(151, 26);
			this.comboBoxSemester.Name = "comboBoxSemester";
			this.comboBoxSemester.Size = new System.Drawing.Size(121, 21);
			this.comboBoxSemester.TabIndex = 1;
			// 
			// semesterBindingSource
			// 
			this.semesterBindingSource.DataMember = "Semester";
			this.semesterBindingSource.DataSource = this.mainDBDataSet;
			// 
			// mainDBDataSet
			// 
			this.mainDBDataSet.DataSetName = "MainDBDataSet";
			this.mainDBDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
			// 
			// comboBoxYear
			// 
			this.comboBoxYear.DataSource = this.studyYearBindingSource;
			this.comboBoxYear.DisplayMember = "StudyYear";
			this.comboBoxYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBoxYear.FormattingEnabled = true;
			this.comboBoxYear.Location = new System.Drawing.Point(12, 26);
			this.comboBoxYear.Name = "comboBoxYear";
			this.comboBoxYear.Size = new System.Drawing.Size(121, 21);
			this.comboBoxYear.TabIndex = 2;
			this.comboBoxYear.ValueMember = "StudyYear";
			// 
			// studyYearBindingSource
			// 
			this.studyYearBindingSource.DataMember = "StudyYear";
			this.studyYearBindingSource.DataSource = this.mainDBDataSetBindingSource;
			// 
			// mainDBDataSetBindingSource
			// 
			this.mainDBDataSetBindingSource.DataSource = this.mainDBDataSet;
			this.mainDBDataSetBindingSource.Position = 0;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(35, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 13);
			this.label1.TabIndex = 3;
			this.label1.Text = "Учебный год";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(190, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(51, 13);
			this.label2.TabIndex = 4;
			this.label2.Text = "Семестр";
			// 
			// studyYearTableAdapter
			// 
			this.studyYearTableAdapter.ClearBeforeFill = true;
			// 
			// semesterTableAdapter
			// 
			this.semesterTableAdapter.ClearBeforeFill = true;
			// 
			// ExportSemester
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(284, 81);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.comboBoxYear);
			this.Controls.Add(this.comboBoxSemester);
			this.Controls.Add(this.buttonExport);
			this.Name = "ExportSemester";
			this.Text = "ExportSemester";
			this.Load += new System.EventHandler(this.ExportSemester_Load);
			((System.ComponentModel.ISupportInitialize)(this.semesterBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSetBindingSource)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonExport;
        private System.Windows.Forms.ComboBox comboBoxSemester;
        private System.Windows.Forms.ComboBox comboBoxYear;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.BindingSource mainDBDataSetBindingSource;
        private MainDBDataSet mainDBDataSet;
        private System.Windows.Forms.BindingSource studyYearBindingSource;
        private MainDBDataSetTableAdapters.StudyYearTableAdapter studyYearTableAdapter;
        private System.Windows.Forms.BindingSource semesterBindingSource;
        private MainDBDataSetTableAdapters.SemesterTableAdapter semesterTableAdapter;
    }
}