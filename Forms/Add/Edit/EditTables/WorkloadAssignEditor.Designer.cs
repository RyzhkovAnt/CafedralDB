namespace CafedralDB.Forms.Add.Edit.EditTables
{
	partial class WorkloadAssignEditor
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
			this.employeeBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.mainDBDataSet = new CafedralDB.MainDBDataSet();
			this.workloadBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.workloadAssignBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.buttonSave = new System.Windows.Forms.Button();
			this.workloadAssignTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.WorkloadAssignTableAdapter();
			this.tableAdapterManager = new CafedralDB.MainDBDataSetTableAdapters.TableAdapterManager();
			this.employeeTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.EmployeeTableAdapter();
			this.workloadTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.WorkloadTableAdapter();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.iDDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.teacherDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
			this.workloadDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
			this.StudentCount = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.Weeks = new System.Windows.Forms.DataGridViewTextBoxColumn();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.employeeBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.workloadBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.workloadAssignBindingSource)).BeginInit();
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
            this.teacherDataGridViewTextBoxColumn,
            this.workloadDataGridViewTextBoxColumn,
            this.StudentCount,
            this.Weeks});
			this.dataGridView1.DataSource = this.workloadAssignBindingSource;
			this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.dataGridView1.Location = new System.Drawing.Point(0, 0);
			this.dataGridView1.Name = "dataGridView1";
			this.dataGridView1.Size = new System.Drawing.Size(401, 214);
			this.dataGridView1.TabIndex = 0;
			// 
			// employeeBindingSource
			// 
			this.employeeBindingSource.DataMember = "Employee";
			this.employeeBindingSource.DataSource = this.mainDBDataSet;
			// 
			// mainDBDataSet
			// 
			this.mainDBDataSet.DataSetName = "MainDBDataSet";
			this.mainDBDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
			// 
			// workloadBindingSource
			// 
			this.workloadBindingSource.DataMember = "Workload";
			this.workloadBindingSource.DataSource = this.mainDBDataSet;
			// 
			// workloadAssignBindingSource
			// 
			this.workloadAssignBindingSource.DataMember = "WorkloadAssign";
			this.workloadAssignBindingSource.DataSource = this.mainDBDataSet;
			// 
			// buttonSave
			// 
			this.buttonSave.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.buttonSave.Location = new System.Drawing.Point(175, -5);
			this.buttonSave.Name = "buttonSave";
			this.buttonSave.Size = new System.Drawing.Size(79, 33);
			this.buttonSave.TabIndex = 1;
			this.buttonSave.Text = "Сохранить";
			this.buttonSave.UseVisualStyleBackColor = true;
			this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
			// 
			// workloadAssignTableAdapter
			// 
			this.workloadAssignTableAdapter.ClearBeforeFill = true;
			// 
			// tableAdapterManager
			// 
			this.tableAdapterManager.AcademicDegreeTableAdapter = null;
			this.tableAdapterManager.AcademicRankTableAdapter = null;
			this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
			this.tableAdapterManager.DepartmentTableAdapter = null;
			this.tableAdapterManager.DisciplineParameterTableAdapter = null;
			this.tableAdapterManager.DisciplineTableAdapter = null;
			this.tableAdapterManager.DisciplineTypeTableAdapter = null;
			this.tableAdapterManager.EmployeeTableAdapter = this.employeeTableAdapter;
			this.tableAdapterManager.FacultyTableAdapter = null;
			this.tableAdapterManager.GroupTableAdapter = null;
			this.tableAdapterManager.NormasTableAdapter = null;
			this.tableAdapterManager.QualificationTableAdapter = null;
			this.tableAdapterManager.SemesterTableAdapter = null;
			this.tableAdapterManager.SpecialityTableAdapter = null;
			this.tableAdapterManager.StudentTableAdapter = null;
			this.tableAdapterManager.StudyFormTableAdapter = null;
			this.tableAdapterManager.StudyYearTableAdapter = null;
			this.tableAdapterManager.UpdateOrder = CafedralDB.MainDBDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
			this.tableAdapterManager.WorkingPositionTableAdapter = null;
			this.tableAdapterManager.WorkloadAssignTableAdapter = this.workloadAssignTableAdapter;
			this.tableAdapterManager.WorkloadTableAdapter = this.workloadTableAdapter;
			// 
			// employeeTableAdapter
			// 
			this.employeeTableAdapter.ClearBeforeFill = true;
			// 
			// workloadTableAdapter
			// 
			this.workloadTableAdapter.ClearBeforeFill = true;
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
			this.splitContainer1.Size = new System.Drawing.Size(401, 243);
			this.splitContainer1.SplitterDistance = 214;
			this.splitContainer1.TabIndex = 2;
			// 
			// iDDataGridViewTextBoxColumn
			// 
			this.iDDataGridViewTextBoxColumn.DataPropertyName = "ID";
			this.iDDataGridViewTextBoxColumn.HeaderText = "ID";
			this.iDDataGridViewTextBoxColumn.Name = "iDDataGridViewTextBoxColumn";
			this.iDDataGridViewTextBoxColumn.Visible = false;
			// 
			// teacherDataGridViewTextBoxColumn
			// 
			this.teacherDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			this.teacherDataGridViewTextBoxColumn.DataPropertyName = "Teacher";
			this.teacherDataGridViewTextBoxColumn.DataSource = this.employeeBindingSource;
			this.teacherDataGridViewTextBoxColumn.DisplayMember = "EmployeeName";
			this.teacherDataGridViewTextBoxColumn.HeaderText = "Teacher";
			this.teacherDataGridViewTextBoxColumn.Name = "teacherDataGridViewTextBoxColumn";
			this.teacherDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
			this.teacherDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
			this.teacherDataGridViewTextBoxColumn.ValueMember = "ID";
			// 
			// workloadDataGridViewTextBoxColumn
			// 
			this.workloadDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			this.workloadDataGridViewTextBoxColumn.DataPropertyName = "Workload";
			this.workloadDataGridViewTextBoxColumn.DataSource = this.workloadBindingSource;
			this.workloadDataGridViewTextBoxColumn.DisplayMember = "Discipline";
			this.workloadDataGridViewTextBoxColumn.HeaderText = "Workload";
			this.workloadDataGridViewTextBoxColumn.Name = "workloadDataGridViewTextBoxColumn";
			this.workloadDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
			this.workloadDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
			this.workloadDataGridViewTextBoxColumn.ValueMember = "ID";
			// 
			// StudentCount
			// 
			this.StudentCount.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			this.StudentCount.DataPropertyName = "StudentCount";
			this.StudentCount.HeaderText = "StudentCount";
			this.StudentCount.Name = "StudentCount";
			// 
			// Weeks
			// 
			this.Weeks.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			this.Weeks.DataPropertyName = "Weeks";
			this.Weeks.HeaderText = "Weeks";
			this.Weeks.Name = "Weeks";
			// 
			// WorkloadAssignEditor
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(401, 243);
			this.Controls.Add(this.splitContainer1);
			this.Name = "WorkloadAssignEditor";
			this.Text = "WorkloadAssignEditor";
			this.Load += new System.EventHandler(this.WorkloadAssignEditor_Load);
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.employeeBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.workloadBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.workloadAssignBindingSource)).EndInit();
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
		private System.Windows.Forms.BindingSource workloadAssignBindingSource;
		private MainDBDataSetTableAdapters.WorkloadAssignTableAdapter workloadAssignTableAdapter;
		private MainDBDataSetTableAdapters.TableAdapterManager tableAdapterManager;
		private MainDBDataSetTableAdapters.EmployeeTableAdapter employeeTableAdapter;
		private System.Windows.Forms.BindingSource employeeBindingSource;
		private MainDBDataSetTableAdapters.WorkloadTableAdapter workloadTableAdapter;
		private System.Windows.Forms.BindingSource workloadBindingSource;
		private System.Windows.Forms.DataGridViewTextBoxColumn iDDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewComboBoxColumn teacherDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewComboBoxColumn workloadDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn StudentCount;
		private System.Windows.Forms.DataGridViewTextBoxColumn Weeks;
		private System.Windows.Forms.SplitContainer splitContainer1;
	}
}