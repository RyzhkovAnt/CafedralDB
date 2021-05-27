namespace CafedralDB.Forms.Add.Edit.EditTables
{
	partial class WorkloadEditor
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
			this.disciplineBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.mainDBDataSet = new CafedralDB.MainDBDataSet();
			this.studyYearBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.semesterBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.groupBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.workloadBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.buttonSave = new System.Windows.Forms.Button();
			this.workloadTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.WorkloadTableAdapter();
			this.tableAdapterManager = new CafedralDB.MainDBDataSetTableAdapters.TableAdapterManager();
			this.disciplineTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.DisciplineTableAdapter();
			this.groupTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.GroupTableAdapter();
			this.semesterTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.SemesterTableAdapter();
			this.studyYearTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.StudyYearTableAdapter();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.iDDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.disciplineDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
			this.studyYearDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
			this.semesterDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
			this.groupDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.disciplineBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.semesterBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.groupBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.workloadBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
			this.splitContainer1.Panel1.SuspendLayout();
			this.splitContainer1.Panel2.SuspendLayout();
			this.splitContainer1.SuspendLayout();
			this.SuspendLayout();
			// 
			// dataGridView1
			// 
			this.dataGridView1.AutoGenerateColumns = false;
			this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.iDDataGridViewTextBoxColumn,
            this.disciplineDataGridViewTextBoxColumn,
            this.studyYearDataGridViewTextBoxColumn,
            this.semesterDataGridViewTextBoxColumn,
            this.groupDataGridViewTextBoxColumn});
			this.dataGridView1.DataSource = this.workloadBindingSource;
			this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.dataGridView1.Location = new System.Drawing.Point(0, 0);
			this.dataGridView1.Name = "dataGridView1";
			this.dataGridView1.Size = new System.Drawing.Size(596, 163);
			this.dataGridView1.TabIndex = 0;
			// 
			// disciplineBindingSource
			// 
			this.disciplineBindingSource.DataMember = "Discipline";
			this.disciplineBindingSource.DataSource = this.mainDBDataSet;
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
			// semesterBindingSource
			// 
			this.semesterBindingSource.DataMember = "Semester";
			this.semesterBindingSource.DataSource = this.mainDBDataSet;
			// 
			// groupBindingSource
			// 
			this.groupBindingSource.DataMember = "Group";
			this.groupBindingSource.DataSource = this.mainDBDataSet;
			// 
			// workloadBindingSource
			// 
			this.workloadBindingSource.DataMember = "Workload";
			this.workloadBindingSource.DataSource = this.mainDBDataSet;
			// 
			// buttonSave
			// 
			this.buttonSave.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.buttonSave.Location = new System.Drawing.Point(274, -2);
			this.buttonSave.Name = "buttonSave";
			this.buttonSave.Size = new System.Drawing.Size(75, 23);
			this.buttonSave.TabIndex = 1;
			this.buttonSave.Text = "Сохранить";
			this.buttonSave.UseVisualStyleBackColor = true;
			this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
			// 
			// workloadTableAdapter
			// 
			this.workloadTableAdapter.ClearBeforeFill = true;
			// 
			// tableAdapterManager
			// 
			this.tableAdapterManager.AcademicDegreeTableAdapter = null;
			this.tableAdapterManager.AcademicRankTableAdapter = null;
			this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
			this.tableAdapterManager.DepartmentTableAdapter = null;
			this.tableAdapterManager.DisciplineParameterTableAdapter = null;
			this.tableAdapterManager.DisciplineTableAdapter = this.disciplineTableAdapter;
			this.tableAdapterManager.DisciplineTypeTableAdapter = null;
			this.tableAdapterManager.EmployeeTableAdapter = null;
			this.tableAdapterManager.FacultyTableAdapter = null;
			this.tableAdapterManager.GroupTableAdapter = this.groupTableAdapter;
			this.tableAdapterManager.NormasTableAdapter = null;
			this.tableAdapterManager.QualificationTableAdapter = null;
			this.tableAdapterManager.SemesterTableAdapter = this.semesterTableAdapter;
			this.tableAdapterManager.SpecialityTableAdapter = null;
			this.tableAdapterManager.StudentTableAdapter = null;
			this.tableAdapterManager.StudyFormTableAdapter = null;
			this.tableAdapterManager.StudyYearTableAdapter = this.studyYearTableAdapter;
			this.tableAdapterManager.UpdateOrder = CafedralDB.MainDBDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
			this.tableAdapterManager.WorkingPositionTableAdapter = null;
			this.tableAdapterManager.WorkloadAssignTableAdapter = null;
			this.tableAdapterManager.WorkloadTableAdapter = this.workloadTableAdapter;
			// 
			// disciplineTableAdapter
			// 
			this.disciplineTableAdapter.ClearBeforeFill = true;
			// 
			// groupTableAdapter
			// 
			this.groupTableAdapter.ClearBeforeFill = true;
			// 
			// semesterTableAdapter
			// 
			this.semesterTableAdapter.ClearBeforeFill = true;
			// 
			// studyYearTableAdapter
			// 
			this.studyYearTableAdapter.ClearBeforeFill = true;
			// 
			// splitContainer1
			// 
			this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.splitContainer1.Location = new System.Drawing.Point(0, 0);
			this.splitContainer1.Name = "splitContainer1";
			this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
			// 
			// splitContainer1.Panel1
			// 
			this.splitContainer1.Panel1.Controls.Add(this.dataGridView1);
			// 
			// splitContainer1.Panel2
			// 
			this.splitContainer1.Panel2.Controls.Add(this.buttonSave);
			this.splitContainer1.Size = new System.Drawing.Size(596, 192);
			this.splitContainer1.SplitterDistance = 163;
			this.splitContainer1.TabIndex = 2;
			// 
			// iDDataGridViewTextBoxColumn
			// 
			this.iDDataGridViewTextBoxColumn.DataPropertyName = "ID";
			this.iDDataGridViewTextBoxColumn.HeaderText = "ID";
			this.iDDataGridViewTextBoxColumn.Name = "iDDataGridViewTextBoxColumn";
			this.iDDataGridViewTextBoxColumn.Visible = false;
			// 
			// disciplineDataGridViewTextBoxColumn
			// 
			this.disciplineDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			this.disciplineDataGridViewTextBoxColumn.DataPropertyName = "Discipline";
			this.disciplineDataGridViewTextBoxColumn.DataSource = this.disciplineBindingSource;
			this.disciplineDataGridViewTextBoxColumn.DisplayMember = "Descr";
			this.disciplineDataGridViewTextBoxColumn.HeaderText = "Discipline";
			this.disciplineDataGridViewTextBoxColumn.MinimumWidth = 250;
			this.disciplineDataGridViewTextBoxColumn.Name = "disciplineDataGridViewTextBoxColumn";
			this.disciplineDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
			this.disciplineDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
			this.disciplineDataGridViewTextBoxColumn.ValueMember = "ID";
			// 
			// studyYearDataGridViewTextBoxColumn
			// 
			this.studyYearDataGridViewTextBoxColumn.DataPropertyName = "StudyYear";
			this.studyYearDataGridViewTextBoxColumn.DataSource = this.studyYearBindingSource;
			this.studyYearDataGridViewTextBoxColumn.DisplayMember = "StudyYear";
			this.studyYearDataGridViewTextBoxColumn.HeaderText = "StudyYear";
			this.studyYearDataGridViewTextBoxColumn.Name = "studyYearDataGridViewTextBoxColumn";
			this.studyYearDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
			this.studyYearDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
			this.studyYearDataGridViewTextBoxColumn.ValueMember = "ID";
			// 
			// semesterDataGridViewTextBoxColumn
			// 
			this.semesterDataGridViewTextBoxColumn.DataPropertyName = "Semester";
			this.semesterDataGridViewTextBoxColumn.DataSource = this.semesterBindingSource;
			this.semesterDataGridViewTextBoxColumn.DisplayMember = "Descr";
			this.semesterDataGridViewTextBoxColumn.HeaderText = "Semester";
			this.semesterDataGridViewTextBoxColumn.Name = "semesterDataGridViewTextBoxColumn";
			this.semesterDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
			this.semesterDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
			this.semesterDataGridViewTextBoxColumn.ValueMember = "ID";
			// 
			// groupDataGridViewTextBoxColumn
			// 
			this.groupDataGridViewTextBoxColumn.DataPropertyName = "Group";
			this.groupDataGridViewTextBoxColumn.DataSource = this.groupBindingSource;
			this.groupDataGridViewTextBoxColumn.DisplayMember = "Descr";
			this.groupDataGridViewTextBoxColumn.HeaderText = "Group";
			this.groupDataGridViewTextBoxColumn.Name = "groupDataGridViewTextBoxColumn";
			this.groupDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
			this.groupDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
			this.groupDataGridViewTextBoxColumn.ValueMember = "ID";
			// 
			// WorkloadEditor
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(596, 192);
			this.Controls.Add(this.splitContainer1);
			this.Name = "WorkloadEditor";
			this.Text = "WorkloadEditor";
			this.Load += new System.EventHandler(this.WorkloadEditor_Load);
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.disciplineBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.semesterBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.groupBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.workloadBindingSource)).EndInit();
			this.splitContainer1.Panel1.ResumeLayout(false);
			this.splitContainer1.Panel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
			this.splitContainer1.ResumeLayout(false);
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.DataGridView dataGridView1;
		private System.Windows.Forms.Button buttonSave;
		private MainDBDataSet mainDBDataSet;
		private System.Windows.Forms.BindingSource workloadBindingSource;
		private MainDBDataSetTableAdapters.WorkloadTableAdapter workloadTableAdapter;
		private MainDBDataSetTableAdapters.TableAdapterManager tableAdapterManager;
		private MainDBDataSetTableAdapters.DisciplineTableAdapter disciplineTableAdapter;
		private System.Windows.Forms.BindingSource disciplineBindingSource;
		private MainDBDataSetTableAdapters.StudyYearTableAdapter studyYearTableAdapter;
		private System.Windows.Forms.BindingSource studyYearBindingSource;
		private MainDBDataSetTableAdapters.SemesterTableAdapter semesterTableAdapter;
		private System.Windows.Forms.BindingSource semesterBindingSource;
		private MainDBDataSetTableAdapters.GroupTableAdapter groupTableAdapter;
		private System.Windows.Forms.BindingSource groupBindingSource;
		private System.Windows.Forms.DataGridViewTextBoxColumn iDDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewComboBoxColumn disciplineDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewComboBoxColumn studyYearDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewComboBoxColumn semesterDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewComboBoxColumn groupDataGridViewTextBoxColumn;
		private System.Windows.Forms.SplitContainer splitContainer1;
	}
}