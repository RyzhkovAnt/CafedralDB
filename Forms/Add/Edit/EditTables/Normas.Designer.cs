namespace CafedralDB.Forms.Add.Edit.EditTables
{
	partial class Normas
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
			this.кодDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.parameterNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.valueDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.commentDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.normasBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.mainDBDataSet = new CafedralDB.MainDBDataSet();
			this.normasTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.NormasTableAdapter();
			this.tableAdapterManager = new CafedralDB.MainDBDataSetTableAdapters.TableAdapterManager();
			this.buttonSave = new System.Windows.Forms.Button();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.normasBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
			this.splitContainer1.Panel1.SuspendLayout();
			this.splitContainer1.Panel2.SuspendLayout();
			this.splitContainer1.SuspendLayout();
			this.SuspendLayout();
			// 
			// dataGridView1
			// 
			this.dataGridView1.AllowUserToAddRows = false;
			this.dataGridView1.AllowUserToDeleteRows = false;
			this.dataGridView1.AllowUserToOrderColumns = true;
			this.dataGridView1.AutoGenerateColumns = false;
			this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.кодDataGridViewTextBoxColumn,
            this.parameterNameDataGridViewTextBoxColumn,
            this.valueDataGridViewTextBoxColumn,
            this.commentDataGridViewTextBoxColumn});
			this.dataGridView1.DataSource = this.normasBindingSource;
			this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.dataGridView1.Location = new System.Drawing.Point(0, 0);
			this.dataGridView1.Name = "dataGridView1";
			this.dataGridView1.Size = new System.Drawing.Size(512, 106);
			this.dataGridView1.TabIndex = 0;
			this.dataGridView1.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellValueChanged);
			// 
			// кодDataGridViewTextBoxColumn
			// 
			this.кодDataGridViewTextBoxColumn.DataPropertyName = "Код";
			this.кодDataGridViewTextBoxColumn.HeaderText = "Код";
			this.кодDataGridViewTextBoxColumn.Name = "кодDataGridViewTextBoxColumn";
			this.кодDataGridViewTextBoxColumn.Visible = false;
			// 
			// parameterNameDataGridViewTextBoxColumn
			// 
			this.parameterNameDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			this.parameterNameDataGridViewTextBoxColumn.DataPropertyName = "ParameterName";
			this.parameterNameDataGridViewTextBoxColumn.HeaderText = "ParameterName";
			this.parameterNameDataGridViewTextBoxColumn.Name = "parameterNameDataGridViewTextBoxColumn";
			this.parameterNameDataGridViewTextBoxColumn.ReadOnly = true;
			// 
			// valueDataGridViewTextBoxColumn
			// 
			this.valueDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			this.valueDataGridViewTextBoxColumn.DataPropertyName = "Value";
			this.valueDataGridViewTextBoxColumn.FillWeight = 30F;
			this.valueDataGridViewTextBoxColumn.HeaderText = "Value";
			this.valueDataGridViewTextBoxColumn.Name = "valueDataGridViewTextBoxColumn";
			// 
			// commentDataGridViewTextBoxColumn
			// 
			this.commentDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			this.commentDataGridViewTextBoxColumn.DataPropertyName = "Comment";
			this.commentDataGridViewTextBoxColumn.FillWeight = 70F;
			this.commentDataGridViewTextBoxColumn.HeaderText = "Comment";
			this.commentDataGridViewTextBoxColumn.Name = "commentDataGridViewTextBoxColumn";
			// 
			// normasBindingSource
			// 
			this.normasBindingSource.DataMember = "Normas";
			this.normasBindingSource.DataSource = this.mainDBDataSet;
			// 
			// mainDBDataSet
			// 
			this.mainDBDataSet.DataSetName = "MainDBDataSet";
			this.mainDBDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
			// 
			// normasTableAdapter
			// 
			this.normasTableAdapter.ClearBeforeFill = true;
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
			this.tableAdapterManager.EmployeeTableAdapter = null;
			this.tableAdapterManager.FacultyTableAdapter = null;
			this.tableAdapterManager.GroupTableAdapter = null;
			this.tableAdapterManager.NormasTableAdapter = this.normasTableAdapter;
			this.tableAdapterManager.QualificationTableAdapter = null;
			this.tableAdapterManager.SemesterTableAdapter = null;
			this.tableAdapterManager.SpecialityTableAdapter = null;
			this.tableAdapterManager.StudentTableAdapter = null;
			this.tableAdapterManager.StudyFormTableAdapter = null;
			this.tableAdapterManager.StudyYearTableAdapter = null;
			this.tableAdapterManager.UpdateOrder = CafedralDB.MainDBDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
			this.tableAdapterManager.WorkingPositionTableAdapter = null;
			this.tableAdapterManager.WorkloadAssignTableAdapter = null;
			this.tableAdapterManager.WorkloadTableAdapter = null;
			// 
			// buttonSave
			// 
			this.buttonSave.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.buttonSave.Location = new System.Drawing.Point(218, -2);
			this.buttonSave.Name = "buttonSave";
			this.buttonSave.Size = new System.Drawing.Size(75, 23);
			this.buttonSave.TabIndex = 1;
			this.buttonSave.Text = "Сохранить";
			this.buttonSave.UseVisualStyleBackColor = true;
			this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
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
			this.splitContainer1.Size = new System.Drawing.Size(512, 135);
			this.splitContainer1.SplitterDistance = 106;
			this.splitContainer1.TabIndex = 2;
			// 
			// Normas
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(512, 135);
			this.Controls.Add(this.splitContainer1);
			this.Name = "Normas";
			this.Text = "Normas";
			this.Load += new System.EventHandler(this.Normas_Load);
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.normasBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).EndInit();
			this.splitContainer1.Panel1.ResumeLayout(false);
			this.splitContainer1.Panel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
			this.splitContainer1.ResumeLayout(false);
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.DataGridView dataGridView1;
		private MainDBDataSet mainDBDataSet;
		private System.Windows.Forms.BindingSource normasBindingSource;
		private MainDBDataSetTableAdapters.NormasTableAdapter normasTableAdapter;
		private MainDBDataSetTableAdapters.TableAdapterManager tableAdapterManager;
		private System.Windows.Forms.Button buttonSave;
		private System.Windows.Forms.DataGridViewTextBoxColumn кодDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn parameterNameDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn valueDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn commentDataGridViewTextBoxColumn;
		private System.Windows.Forms.SplitContainer splitContainer1;
	}
}