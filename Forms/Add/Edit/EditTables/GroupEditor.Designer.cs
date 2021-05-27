namespace CafedralDB.Forms.Add.Edit.EditTables
{
	partial class GroupEditor
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
            this.specialityBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.mainDBDataSet = new CafedralDB.MainDBDataSet();
            this.qualificationBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.studyFormBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.studyYearBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.groupBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.buttonSave = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.groupTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.GroupTableAdapter();
            this.specialityTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.SpecialityTableAdapter();
            this.qualificationTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.QualificationTableAdapter();
            this.studyFormTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.StudyFormTableAdapter();
            this.studyYearTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.StudyYearTableAdapter();
            this.iDDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.descrDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.specialityDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.qualificationDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.studyFormDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.studentCountDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.entryYearDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.SubgroupCount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.specialityBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.qualificationBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.studyFormBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupBindingSource)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.iDDataGridViewTextBoxColumn,
            this.descrDataGridViewTextBoxColumn,
            this.specialityDataGridViewTextBoxColumn,
            this.qualificationDataGridViewTextBoxColumn,
            this.studyFormDataGridViewTextBoxColumn,
            this.studentCountDataGridViewTextBoxColumn,
            this.entryYearDataGridViewTextBoxColumn,
            this.SubgroupCount});
            this.dataGridView1.DataSource = this.groupBindingSource;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(646, 195);
            this.dataGridView1.TabIndex = 0;
            // 
            // specialityBindingSource
            // 
            this.specialityBindingSource.DataMember = "Speciality";
            this.specialityBindingSource.DataSource = this.mainDBDataSet;
            // 
            // mainDBDataSet
            // 
            this.mainDBDataSet.DataSetName = "MainDBDataSet";
            this.mainDBDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // qualificationBindingSource
            // 
            this.qualificationBindingSource.DataMember = "Qualification";
            this.qualificationBindingSource.DataSource = this.mainDBDataSet;
            // 
            // studyFormBindingSource
            // 
            this.studyFormBindingSource.DataMember = "StudyForm";
            this.studyFormBindingSource.DataSource = this.mainDBDataSet;
            // 
            // studyYearBindingSource
            // 
            this.studyYearBindingSource.DataMember = "StudyYear";
            this.studyYearBindingSource.DataSource = this.mainDBDataSet;
            // 
            // groupBindingSource
            // 
            this.groupBindingSource.DataMember = "Group";
            this.groupBindingSource.DataSource = this.mainDBDataSet;
            // 
            // buttonSave
            // 
            this.buttonSave.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.buttonSave.Location = new System.Drawing.Point(275, 3);
            this.buttonSave.MaximumSize = new System.Drawing.Size(100, 100);
            this.buttonSave.MinimumSize = new System.Drawing.Size(100, 20);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(100, 28);
            this.buttonSave.TabIndex = 1;
            this.buttonSave.Text = "Сохранить";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
            // 
            // panel1
            // 
            this.panel1.AutoSize = true;
            this.panel1.Controls.Add(this.dataGridView1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(646, 195);
            this.panel1.TabIndex = 2;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.buttonSave);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 161);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(646, 34);
            this.panel2.TabIndex = 3;
            // 
            // groupTableAdapter
            // 
            this.groupTableAdapter.ClearBeforeFill = true;
            // 
            // specialityTableAdapter
            // 
            this.specialityTableAdapter.ClearBeforeFill = true;
            // 
            // qualificationTableAdapter
            // 
            this.qualificationTableAdapter.ClearBeforeFill = true;
            // 
            // studyFormTableAdapter
            // 
            this.studyFormTableAdapter.ClearBeforeFill = true;
            // 
            // studyYearTableAdapter
            // 
            this.studyYearTableAdapter.ClearBeforeFill = true;
            // 
            // iDDataGridViewTextBoxColumn
            // 
            this.iDDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.iDDataGridViewTextBoxColumn.DataPropertyName = "ID";
            this.iDDataGridViewTextBoxColumn.FillWeight = 10F;
            this.iDDataGridViewTextBoxColumn.HeaderText = "ID";
            this.iDDataGridViewTextBoxColumn.Name = "iDDataGridViewTextBoxColumn";
            // 
            // descrDataGridViewTextBoxColumn
            // 
            this.descrDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.descrDataGridViewTextBoxColumn.DataPropertyName = "Descr";
            this.descrDataGridViewTextBoxColumn.FillWeight = 30F;
            this.descrDataGridViewTextBoxColumn.HeaderText = "Descr";
            this.descrDataGridViewTextBoxColumn.Name = "descrDataGridViewTextBoxColumn";
            // 
            // specialityDataGridViewTextBoxColumn
            // 
            this.specialityDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.specialityDataGridViewTextBoxColumn.DataPropertyName = "Speciality";
            this.specialityDataGridViewTextBoxColumn.DataSource = this.specialityBindingSource;
            this.specialityDataGridViewTextBoxColumn.DisplayMember = "Descr";
            this.specialityDataGridViewTextBoxColumn.FillWeight = 20F;
            this.specialityDataGridViewTextBoxColumn.HeaderText = "Speciality";
            this.specialityDataGridViewTextBoxColumn.Name = "specialityDataGridViewTextBoxColumn";
            this.specialityDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.specialityDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.specialityDataGridViewTextBoxColumn.ValueMember = "ID";
            // 
            // qualificationDataGridViewTextBoxColumn
            // 
            this.qualificationDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.qualificationDataGridViewTextBoxColumn.DataPropertyName = "Qualification";
            this.qualificationDataGridViewTextBoxColumn.DataSource = this.qualificationBindingSource;
            this.qualificationDataGridViewTextBoxColumn.DisplayMember = "Descr";
            this.qualificationDataGridViewTextBoxColumn.FillWeight = 30F;
            this.qualificationDataGridViewTextBoxColumn.HeaderText = "Qualification";
            this.qualificationDataGridViewTextBoxColumn.Name = "qualificationDataGridViewTextBoxColumn";
            this.qualificationDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.qualificationDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.qualificationDataGridViewTextBoxColumn.ValueMember = "ID";
            // 
            // studyFormDataGridViewTextBoxColumn
            // 
            this.studyFormDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.studyFormDataGridViewTextBoxColumn.DataPropertyName = "StudyForm";
            this.studyFormDataGridViewTextBoxColumn.DataSource = this.studyFormBindingSource;
            this.studyFormDataGridViewTextBoxColumn.DisplayMember = "DescrRus";
            this.studyFormDataGridViewTextBoxColumn.FillWeight = 20F;
            this.studyFormDataGridViewTextBoxColumn.HeaderText = "StudyForm";
            this.studyFormDataGridViewTextBoxColumn.Name = "studyFormDataGridViewTextBoxColumn";
            this.studyFormDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.studyFormDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.studyFormDataGridViewTextBoxColumn.ValueMember = "ID";
            // 
            // studentCountDataGridViewTextBoxColumn
            // 
            this.studentCountDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.studentCountDataGridViewTextBoxColumn.DataPropertyName = "StudentCount";
            this.studentCountDataGridViewTextBoxColumn.FillWeight = 10F;
            this.studentCountDataGridViewTextBoxColumn.HeaderText = "StudentCount";
            this.studentCountDataGridViewTextBoxColumn.Name = "studentCountDataGridViewTextBoxColumn";
            // 
            // entryYearDataGridViewTextBoxColumn
            // 
            this.entryYearDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.entryYearDataGridViewTextBoxColumn.DataPropertyName = "EntryYear";
            this.entryYearDataGridViewTextBoxColumn.DataSource = this.studyYearBindingSource;
            this.entryYearDataGridViewTextBoxColumn.DisplayMember = "StudyYear";
            this.entryYearDataGridViewTextBoxColumn.FillWeight = 20F;
            this.entryYearDataGridViewTextBoxColumn.HeaderText = "EntryYear";
            this.entryYearDataGridViewTextBoxColumn.Name = "entryYearDataGridViewTextBoxColumn";
            this.entryYearDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.entryYearDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.entryYearDataGridViewTextBoxColumn.ValueMember = "ID";
            // 
            // SubgroupCount
            // 
            this.SubgroupCount.DataPropertyName = "SubgroupCount";
            this.SubgroupCount.HeaderText = "SubgroupCount";
            this.SubgroupCount.Name = "SubgroupCount";
            // 
            // GroupEditor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(646, 195);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "GroupEditor";
            this.Text = "GroupEditor";
            this.Load += new System.EventHandler(this.GroupEditor_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.specialityBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.qualificationBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.studyFormBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupBindingSource)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.DataGridView dataGridView1;
		private System.Windows.Forms.Button buttonSave;
		private MainDBDataSet mainDBDataSet;
		private System.Windows.Forms.BindingSource groupBindingSource;
		private MainDBDataSetTableAdapters.GroupTableAdapter groupTableAdapter;
		private System.Windows.Forms.BindingSource specialityBindingSource;
		private MainDBDataSetTableAdapters.SpecialityTableAdapter specialityTableAdapter;
		private System.Windows.Forms.BindingSource qualificationBindingSource;
		private MainDBDataSetTableAdapters.QualificationTableAdapter qualificationTableAdapter;
		private System.Windows.Forms.BindingSource studyFormBindingSource;
		private MainDBDataSetTableAdapters.StudyFormTableAdapter studyFormTableAdapter;
		private System.Windows.Forms.BindingSource studyYearBindingSource;
		private MainDBDataSetTableAdapters.StudyYearTableAdapter studyYearTableAdapter;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn iDDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn descrDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewComboBoxColumn specialityDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewComboBoxColumn qualificationDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewComboBoxColumn studyFormDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn studentCountDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewComboBoxColumn entryYearDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn SubgroupCount;
    }
}