namespace CafedralDB.Forms.Add.Assign
{
    partial class AssignTeachersSimpleForm
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.iDDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.дисциплинаDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.группаDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.формаDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.семестрDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.StudyYear = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Cost = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Assigned = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.dataTableForAssignsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.mainDBDataSet = new CafedralDB.MainDBDataSet();
            this.studyYearBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.labelAllTeachersHours = new System.Windows.Forms.Label();
            this.buttonAssign = new System.Windows.Forms.Button();
            this.buttonFree = new System.Windows.Forms.Button();
            this.comboBoxStudyYear = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxWorkloadHours = new System.Windows.Forms.TextBox();
            this.studyYearTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.StudyYearTableAdapter();
            this.dataTableForAssignsTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.DataTableForAssignsTableAdapter();
            this.dataGridViewFreeTeachers = new System.Windows.Forms.DataGridView();
            this.ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TeacherName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewAssignedTeachers = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewEmployeeSumms = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Summ = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.textBoxStudentCount = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTableForAssignsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewFreeTeachers)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewAssignedTeachers)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEmployeeSumms)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.iDDataGridViewTextBoxColumn,
            this.дисциплинаDataGridViewTextBoxColumn,
            this.группаDataGridViewTextBoxColumn,
            this.формаDataGridViewTextBoxColumn,
            this.семестрDataGridViewTextBoxColumn,
            this.StudyYear,
            this.Cost,
            this.Assigned});
            this.dataGridView1.DataSource = this.dataTableForAssignsBindingSource;
            this.dataGridView1.Location = new System.Drawing.Point(12, 33);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.ShowEditingIcon = false;
            this.dataGridView1.Size = new System.Drawing.Size(807, 211);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.dataGridView1_DataBindingComplete);
            this.dataGridView1.SelectionChanged += new System.EventHandler(this.dataGridView1_SelectionChanged);
            // 
            // iDDataGridViewTextBoxColumn
            // 
            this.iDDataGridViewTextBoxColumn.DataPropertyName = "ID";
            this.iDDataGridViewTextBoxColumn.HeaderText = "ID";
            this.iDDataGridViewTextBoxColumn.Name = "iDDataGridViewTextBoxColumn";
            this.iDDataGridViewTextBoxColumn.ReadOnly = true;
            this.iDDataGridViewTextBoxColumn.Visible = false;
            // 
            // дисциплинаDataGridViewTextBoxColumn
            // 
            this.дисциплинаDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.дисциплинаDataGridViewTextBoxColumn.DataPropertyName = "Дисциплина";
            this.дисциплинаDataGridViewTextBoxColumn.FillWeight = 35F;
            this.дисциплинаDataGridViewTextBoxColumn.HeaderText = "Дисциплина";
            this.дисциплинаDataGridViewTextBoxColumn.Name = "дисциплинаDataGridViewTextBoxColumn";
            this.дисциплинаDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // группаDataGridViewTextBoxColumn
            // 
            this.группаDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.группаDataGridViewTextBoxColumn.DataPropertyName = "Группа";
            this.группаDataGridViewTextBoxColumn.FillWeight = 20F;
            this.группаDataGridViewTextBoxColumn.HeaderText = "Группа";
            this.группаDataGridViewTextBoxColumn.Name = "группаDataGridViewTextBoxColumn";
            this.группаDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // формаDataGridViewTextBoxColumn
            // 
            this.формаDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.формаDataGridViewTextBoxColumn.DataPropertyName = "Форма";
            this.формаDataGridViewTextBoxColumn.FillWeight = 15F;
            this.формаDataGridViewTextBoxColumn.HeaderText = "Форма";
            this.формаDataGridViewTextBoxColumn.Name = "формаDataGridViewTextBoxColumn";
            this.формаDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // семестрDataGridViewTextBoxColumn
            // 
            this.семестрDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.семестрDataGridViewTextBoxColumn.DataPropertyName = "Семестр";
            this.семестрDataGridViewTextBoxColumn.FillWeight = 10F;
            this.семестрDataGridViewTextBoxColumn.HeaderText = "Семестр";
            this.семестрDataGridViewTextBoxColumn.Name = "семестрDataGridViewTextBoxColumn";
            this.семестрDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // StudyYear
            // 
            this.StudyYear.DataPropertyName = "StudyYear";
            this.StudyYear.HeaderText = "StudyYear";
            this.StudyYear.Name = "StudyYear";
            this.StudyYear.ReadOnly = true;
            this.StudyYear.Visible = false;
            // 
            // Cost
            // 
            this.Cost.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Cost.FillWeight = 5F;
            this.Cost.HeaderText = "Часов";
            this.Cost.Name = "Cost";
            this.Cost.ReadOnly = true;
            // 
            // Assigned
            // 
            this.Assigned.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Assigned.FillWeight = 5F;
            this.Assigned.HeaderText = "Назн";
            this.Assigned.Name = "Assigned";
            this.Assigned.ReadOnly = true;
            // 
            // dataTableForAssignsBindingSource
            // 
            this.dataTableForAssignsBindingSource.DataMember = "DataTableForAssigns";
            this.dataTableForAssignsBindingSource.DataSource = this.mainDBDataSet;
            // 
            // mainDBDataSet
            // 
            this.mainDBDataSet.DataSetName = "MainDBDataSet";
            this.mainDBDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // studyYearBindingSource
            // 
            this.studyYearBindingSource.DataMember = "StudyYear";
            this.studyYearBindingSource.DataSource = this.mainDBDataSet;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 247);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(144, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Свободные преподаватели";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(332, 247);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(156, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Назначенные преподаватели";
            // 
            // labelAllTeachersHours
            // 
            this.labelAllTeachersHours.AutoSize = true;
            this.labelAllTeachersHours.Location = new System.Drawing.Point(656, 247);
            this.labelAllTeachersHours.Name = "labelAllTeachersHours";
            this.labelAllTeachersHours.Size = new System.Drawing.Size(163, 13);
            this.labelAllTeachersHours.TabIndex = 5;
            this.labelAllTeachersHours.Text = "Всего часов у преподавателей";
            // 
            // buttonAssign
            // 
            this.buttonAssign.Location = new System.Drawing.Point(209, 279);
            this.buttonAssign.Name = "buttonAssign";
            this.buttonAssign.Size = new System.Drawing.Size(75, 23);
            this.buttonAssign.TabIndex = 7;
            this.buttonAssign.Text = "=>";
            this.buttonAssign.UseVisualStyleBackColor = true;
            this.buttonAssign.Click += new System.EventHandler(this.buttonAssign_Click);
            // 
            // buttonFree
            // 
            this.buttonFree.Location = new System.Drawing.Point(209, 362);
            this.buttonFree.Name = "buttonFree";
            this.buttonFree.Size = new System.Drawing.Size(75, 23);
            this.buttonFree.TabIndex = 8;
            this.buttonFree.Text = "<=";
            this.buttonFree.UseVisualStyleBackColor = true;
            this.buttonFree.Click += new System.EventHandler(this.buttonFree_Click);
            // 
            // comboBoxStudyYear
            // 
            this.comboBoxStudyYear.DataSource = this.studyYearBindingSource;
            this.comboBoxStudyYear.DisplayMember = "StudyYear";
            this.comboBoxStudyYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxStudyYear.FormattingEnabled = true;
            this.comboBoxStudyYear.Location = new System.Drawing.Point(87, 6);
            this.comboBoxStudyYear.Name = "comboBoxStudyYear";
            this.comboBoxStudyYear.Size = new System.Drawing.Size(121, 21);
            this.comboBoxStudyYear.TabIndex = 9;
            this.comboBoxStudyYear.ValueMember = "ID";
            this.comboBoxStudyYear.SelectionChangeCommitted += new System.EventHandler(this.comboBoxStudyYear_SelectionChangeCommitted);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Учебный год";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(162, 311);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(39, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Часов";
            // 
            // textBoxWorkloadHours
            // 
            this.textBoxWorkloadHours.Location = new System.Drawing.Point(209, 308);
            this.textBoxWorkloadHours.Name = "textBoxWorkloadHours";
            this.textBoxWorkloadHours.Size = new System.Drawing.Size(100, 20);
            this.textBoxWorkloadHours.TabIndex = 12;
            // 
            // studyYearTableAdapter
            // 
            this.studyYearTableAdapter.ClearBeforeFill = true;
            // 
            // dataTableForAssignsTableAdapter
            // 
            this.dataTableForAssignsTableAdapter.ClearBeforeFill = true;
            // 
            // dataGridViewFreeTeachers
            // 
            this.dataGridViewFreeTeachers.AllowUserToAddRows = false;
            this.dataGridViewFreeTeachers.AllowUserToDeleteRows = false;
            this.dataGridViewFreeTeachers.AllowUserToResizeColumns = false;
            this.dataGridViewFreeTeachers.AllowUserToResizeRows = false;
            this.dataGridViewFreeTeachers.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridViewFreeTeachers.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGridViewFreeTeachers.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridViewFreeTeachers.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dataGridViewFreeTeachers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewFreeTeachers.ColumnHeadersVisible = false;
            this.dataGridViewFreeTeachers.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ID,
            this.TeacherName});
            this.dataGridViewFreeTeachers.Location = new System.Drawing.Point(15, 263);
            this.dataGridViewFreeTeachers.MultiSelect = false;
            this.dataGridViewFreeTeachers.Name = "dataGridViewFreeTeachers";
            this.dataGridViewFreeTeachers.ReadOnly = true;
            this.dataGridViewFreeTeachers.RowHeadersVisible = false;
            this.dataGridViewFreeTeachers.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewFreeTeachers.ShowEditingIcon = false;
            this.dataGridViewFreeTeachers.Size = new System.Drawing.Size(144, 150);
            this.dataGridViewFreeTeachers.TabIndex = 13;
            // 
            // ID
            // 
            this.ID.HeaderText = "ID";
            this.ID.Name = "ID";
            this.ID.ReadOnly = true;
            this.ID.Visible = false;
            this.ID.Width = 72;
            // 
            // TeacherName
            // 
            this.TeacherName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.TeacherName.HeaderText = "TeacherName";
            this.TeacherName.Name = "TeacherName";
            this.TeacherName.ReadOnly = true;
            // 
            // dataGridViewAssignedTeachers
            // 
            this.dataGridViewAssignedTeachers.AllowUserToAddRows = false;
            this.dataGridViewAssignedTeachers.AllowUserToDeleteRows = false;
            this.dataGridViewAssignedTeachers.AllowUserToResizeColumns = false;
            this.dataGridViewAssignedTeachers.AllowUserToResizeRows = false;
            this.dataGridViewAssignedTeachers.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridViewAssignedTeachers.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGridViewAssignedTeachers.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridViewAssignedTeachers.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dataGridViewAssignedTeachers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewAssignedTeachers.ColumnHeadersVisible = false;
            this.dataGridViewAssignedTeachers.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2});
            this.dataGridViewAssignedTeachers.Location = new System.Drawing.Point(335, 263);
            this.dataGridViewAssignedTeachers.MultiSelect = false;
            this.dataGridViewAssignedTeachers.Name = "dataGridViewAssignedTeachers";
            this.dataGridViewAssignedTeachers.ReadOnly = true;
            this.dataGridViewAssignedTeachers.RowHeadersVisible = false;
            this.dataGridViewAssignedTeachers.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewAssignedTeachers.ShowEditingIcon = false;
            this.dataGridViewAssignedTeachers.Size = new System.Drawing.Size(153, 150);
            this.dataGridViewAssignedTeachers.TabIndex = 14;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.HeaderText = "ID";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Visible = false;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn2.HeaderText = "TeacherName";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            // 
            // dataGridViewEmployeeSumms
            // 
            this.dataGridViewEmployeeSumms.AllowUserToAddRows = false;
            this.dataGridViewEmployeeSumms.AllowUserToDeleteRows = false;
            this.dataGridViewEmployeeSumms.AllowUserToResizeColumns = false;
            this.dataGridViewEmployeeSumms.AllowUserToResizeRows = false;
            this.dataGridViewEmployeeSumms.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridViewEmployeeSumms.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGridViewEmployeeSumms.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridViewEmployeeSumms.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewEmployeeSumms.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn4,
            this.Summ});
            this.dataGridViewEmployeeSumms.Location = new System.Drawing.Point(659, 263);
            this.dataGridViewEmployeeSumms.MultiSelect = false;
            this.dataGridViewEmployeeSumms.Name = "dataGridViewEmployeeSumms";
            this.dataGridViewEmployeeSumms.ReadOnly = true;
            this.dataGridViewEmployeeSumms.RowHeadersVisible = false;
            this.dataGridViewEmployeeSumms.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewEmployeeSumms.ShowEditingIcon = false;
            this.dataGridViewEmployeeSumms.ShowRowErrors = false;
            this.dataGridViewEmployeeSumms.Size = new System.Drawing.Size(160, 150);
            this.dataGridViewEmployeeSumms.TabIndex = 15;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn4.HeaderText = "Имя";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.ReadOnly = true;
            // 
            // Summ
            // 
            this.Summ.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Summ.FillWeight = 50F;
            this.Summ.HeaderText = "Сумма";
            this.Summ.Name = "Summ";
            this.Summ.ReadOnly = true;
            // 
            // textBoxStudentCount
            // 
            this.textBoxStudentCount.Location = new System.Drawing.Point(227, 336);
            this.textBoxStudentCount.Name = "textBoxStudentCount";
            this.textBoxStudentCount.Size = new System.Drawing.Size(36, 20);
            this.textBoxStudentCount.TabIndex = 17;
            this.textBoxStudentCount.Text = "1";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(162, 339);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 13);
            this.label5.TabIndex = 16;
            this.label5.Text = "Студентов";
            // 
            // AssignTeachersSimpleForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(839, 424);
            this.Controls.Add(this.textBoxStudentCount);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.dataGridViewEmployeeSumms);
            this.Controls.Add(this.dataGridViewAssignedTeachers);
            this.Controls.Add(this.dataGridViewFreeTeachers);
            this.Controls.Add(this.textBoxWorkloadHours);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.comboBoxStudyYear);
            this.Controls.Add(this.buttonFree);
            this.Controls.Add(this.buttonAssign);
            this.Controls.Add(this.labelAllTeachersHours);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "AssignTeachersSimpleForm";
            this.Text = "AssignTeachersSimpleForm";
            this.Load += new System.EventHandler(this.AssignTeachersSimpleForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTableForAssignsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewFreeTeachers)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewAssignedTeachers)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEmployeeSumms)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label labelAllTeachersHours;
        private System.Windows.Forms.Button buttonAssign;
        private System.Windows.Forms.Button buttonFree;
        private System.Windows.Forms.ComboBox comboBoxStudyYear;
        private System.Windows.Forms.Label label3;

		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox textBoxWorkloadHours;
		private System.Windows.Forms.DataGridViewTextBoxColumn descrDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn descrRusDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn groupDescrDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn semestrDescrDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn yearIDDataGridViewTextBoxColumn;
		private MainDBDataSet mainDBDataSet;
		private System.Windows.Forms.BindingSource studyYearBindingSource;
		private MainDBDataSetTableAdapters.StudyYearTableAdapter studyYearTableAdapter;
		private System.Windows.Forms.BindingSource dataTableForAssignsBindingSource;
		private MainDBDataSetTableAdapters.DataTableForAssignsTableAdapter dataTableForAssignsTableAdapter;
		private System.Windows.Forms.DataGridView dataGridViewFreeTeachers;
		private System.Windows.Forms.DataGridView dataGridViewAssignedTeachers;
		private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
		private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
		private System.Windows.Forms.DataGridViewTextBoxColumn ID;
		private System.Windows.Forms.DataGridViewTextBoxColumn TeacherName;
		private System.Windows.Forms.DataGridView dataGridViewEmployeeSumms;
		private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
		private System.Windows.Forms.DataGridViewTextBoxColumn Summ;
		private System.Windows.Forms.DataGridViewTextBoxColumn iDDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn дисциплинаDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn группаDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn формаDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn семестрDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn StudyYear;
		private System.Windows.Forms.DataGridViewTextBoxColumn Cost;
		private System.Windows.Forms.DataGridViewCheckBoxColumn Assigned;
		private System.Windows.Forms.TextBox textBoxStudentCount;
		private System.Windows.Forms.Label label5;
	}
}