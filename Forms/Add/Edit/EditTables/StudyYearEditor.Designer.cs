namespace CafedralDB.Forms.Add.Edit.EditTables
{
	partial class StudyYearEditor
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
			this.buttonSave = new System.Windows.Forms.Button();
			this.dataGridViewStudyYear = new System.Windows.Forms.DataGridView();
			this.studyYearBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.mainDBDataSet = new CafedralDB.MainDBDataSet();
			this.studyYearTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.StudyYearTableAdapter();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.iDDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.studyYearDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
			((System.ComponentModel.ISupportInitialize)(this.dataGridViewStudyYear)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
			this.splitContainer1.Panel1.SuspendLayout();
			this.splitContainer1.Panel2.SuspendLayout();
			this.splitContainer1.SuspendLayout();
			this.SuspendLayout();
			// 
			// buttonSave
			// 
			this.buttonSave.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.buttonSave.Location = new System.Drawing.Point(80, -2);
			this.buttonSave.Name = "buttonSave";
			this.buttonSave.Size = new System.Drawing.Size(75, 23);
			this.buttonSave.TabIndex = 3;
			this.buttonSave.Text = "Сохранить";
			this.buttonSave.UseVisualStyleBackColor = true;
			this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
			// 
			// dataGridViewStudyYear
			// 
			this.dataGridViewStudyYear.AutoGenerateColumns = false;
			this.dataGridViewStudyYear.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridViewStudyYear.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.iDDataGridViewTextBoxColumn,
            this.studyYearDataGridViewTextBoxColumn});
			this.dataGridViewStudyYear.DataSource = this.studyYearBindingSource;
			this.dataGridViewStudyYear.Dock = System.Windows.Forms.DockStyle.Fill;
			this.dataGridViewStudyYear.Location = new System.Drawing.Point(0, 0);
			this.dataGridViewStudyYear.Name = "dataGridViewStudyYear";
			this.dataGridViewStudyYear.Size = new System.Drawing.Size(227, 160);
			this.dataGridViewStudyYear.TabIndex = 2;
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
			this.splitContainer1.Panel1.Controls.Add(this.dataGridViewStudyYear);
			// 
			// splitContainer1.Panel2
			// 
			this.splitContainer1.Panel2.Controls.Add(this.buttonSave);
			this.splitContainer1.Size = new System.Drawing.Size(227, 189);
			this.splitContainer1.SplitterDistance = 160;
			this.splitContainer1.TabIndex = 4;
			// 
			// iDDataGridViewTextBoxColumn
			// 
			this.iDDataGridViewTextBoxColumn.DataPropertyName = "ID";
			this.iDDataGridViewTextBoxColumn.HeaderText = "ID";
			this.iDDataGridViewTextBoxColumn.Name = "iDDataGridViewTextBoxColumn";
			this.iDDataGridViewTextBoxColumn.Visible = false;
			// 
			// studyYearDataGridViewTextBoxColumn
			// 
			this.studyYearDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			this.studyYearDataGridViewTextBoxColumn.DataPropertyName = "StudyYear";
			this.studyYearDataGridViewTextBoxColumn.HeaderText = "StudyYear";
			this.studyYearDataGridViewTextBoxColumn.Name = "studyYearDataGridViewTextBoxColumn";
			// 
			// StudyYearEditor
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(227, 189);
			this.Controls.Add(this.splitContainer1);
			this.Name = "StudyYearEditor";
			this.Text = "StudyYearEditor";
			this.Load += new System.EventHandler(this.StudyYearEditor_Load);
			((System.ComponentModel.ISupportInitialize)(this.dataGridViewStudyYear)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).EndInit();
			this.splitContainer1.Panel1.ResumeLayout(false);
			this.splitContainer1.Panel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
			this.splitContainer1.ResumeLayout(false);
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button buttonSave;
		private System.Windows.Forms.DataGridView dataGridViewStudyYear;
		private MainDBDataSet mainDBDataSet;
		private System.Windows.Forms.BindingSource studyYearBindingSource;
		private MainDBDataSetTableAdapters.StudyYearTableAdapter studyYearTableAdapter;
		private System.Windows.Forms.DataGridViewTextBoxColumn iDDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn studyYearDataGridViewTextBoxColumn;
		private System.Windows.Forms.SplitContainer splitContainer1;
	}
}