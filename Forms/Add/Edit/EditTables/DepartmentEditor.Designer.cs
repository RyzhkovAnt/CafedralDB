namespace CafedralDB.Forms.Add.Edit.EditTables
{
	partial class DepartmentEditor
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
			this.departmentBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.mainDBDataSet = new CafedralDB.MainDBDataSet();
			this.buttonSave = new System.Windows.Forms.Button();
			this.departmentTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.DepartmentTableAdapter();
			this.iDDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.codeNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.descrDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.departmentBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).BeginInit();
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
            this.codeNumberDataGridViewTextBoxColumn,
            this.descrDataGridViewTextBoxColumn});
			this.dataGridView1.DataSource = this.departmentBindingSource;
			this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.dataGridView1.Location = new System.Drawing.Point(0, 0);
			this.dataGridView1.Name = "dataGridView1";
			this.dataGridView1.Size = new System.Drawing.Size(357, 211);
			this.dataGridView1.TabIndex = 0;
			// 
			// departmentBindingSource
			// 
			this.departmentBindingSource.DataMember = "Department";
			this.departmentBindingSource.DataSource = this.mainDBDataSet;
			// 
			// mainDBDataSet
			// 
			this.mainDBDataSet.DataSetName = "MainDBDataSet";
			this.mainDBDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
			// 
			// buttonSave
			// 
			this.buttonSave.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.buttonSave.Location = new System.Drawing.Point(141, 1);
			this.buttonSave.Name = "buttonSave";
			this.buttonSave.Size = new System.Drawing.Size(75, 23);
			this.buttonSave.TabIndex = 1;
			this.buttonSave.Text = "Сохранить";
			this.buttonSave.UseVisualStyleBackColor = true;
			this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
			// 
			// departmentTableAdapter
			// 
			this.departmentTableAdapter.ClearBeforeFill = true;
			// 
			// iDDataGridViewTextBoxColumn
			// 
			this.iDDataGridViewTextBoxColumn.DataPropertyName = "ID";
			this.iDDataGridViewTextBoxColumn.HeaderText = "ID";
			this.iDDataGridViewTextBoxColumn.Name = "iDDataGridViewTextBoxColumn";
			this.iDDataGridViewTextBoxColumn.Visible = false;
			this.iDDataGridViewTextBoxColumn.Width = 30;
			// 
			// codeNumberDataGridViewTextBoxColumn
			// 
			this.codeNumberDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
			this.codeNumberDataGridViewTextBoxColumn.DataPropertyName = "CodeNumber";
			this.codeNumberDataGridViewTextBoxColumn.HeaderText = "CodeNumber";
			this.codeNumberDataGridViewTextBoxColumn.Name = "codeNumberDataGridViewTextBoxColumn";
			this.codeNumberDataGridViewTextBoxColumn.Width = 94;
			// 
			// descrDataGridViewTextBoxColumn
			// 
			this.descrDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			this.descrDataGridViewTextBoxColumn.DataPropertyName = "Descr";
			this.descrDataGridViewTextBoxColumn.HeaderText = "Descr";
			this.descrDataGridViewTextBoxColumn.Name = "descrDataGridViewTextBoxColumn";
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
			this.splitContainer1.Size = new System.Drawing.Size(357, 240);
			this.splitContainer1.SplitterDistance = 211;
			this.splitContainer1.TabIndex = 2;
			// 
			// DepartmentEditor
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(357, 240);
			this.Controls.Add(this.splitContainer1);
			this.Name = "DepartmentEditor";
			this.Text = "DepartmentEditor";
			this.Load += new System.EventHandler(this.DepartmentEditor_Load);
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.departmentBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).EndInit();
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
		private System.Windows.Forms.BindingSource departmentBindingSource;
		private MainDBDataSetTableAdapters.DepartmentTableAdapter departmentTableAdapter;
		private System.Windows.Forms.DataGridViewTextBoxColumn iDDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn codeNumberDataGridViewTextBoxColumn;
		private System.Windows.Forms.DataGridViewTextBoxColumn descrDataGridViewTextBoxColumn;
		private System.Windows.Forms.SplitContainer splitContainer1;
	}
}