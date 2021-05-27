namespace CafedralDB.Forms.Add.Import
{
    partial class ImportForm
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
			this.comboBoxYear = new System.Windows.Forms.ComboBox();
			this.studyYearBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.mainDBDataSet = new CafedralDB.MainDBDataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.buttonImport = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.listBox1 = new System.Windows.Forms.ListBox();
			this.yearsQueryBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.mainDBDataSet1 = new CafedralDB.MainDBDataSet();
			this.yearsExistBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.studyYearTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.StudyYearTableAdapter();
			this.yearsExistTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.YearsExistQueryTableAdapter();
			this.yearsQueryTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.YearsQueryTableAdapter();
			this.yearsExistQueryBindingSource = new System.Windows.Forms.BindingSource(this.components);
			((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.yearsQueryBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.yearsExistBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.yearsExistQueryBindingSource)).BeginInit();
			this.SuspendLayout();
			// 
			// comboBoxYear
			// 
			this.comboBoxYear.DataSource = this.studyYearBindingSource;
			this.comboBoxYear.DisplayMember = "StudyYear";
			this.comboBoxYear.FormattingEnabled = true;
			this.comboBoxYear.Location = new System.Drawing.Point(12, 26);
			this.comboBoxYear.Name = "comboBoxYear";
			this.comboBoxYear.Size = new System.Drawing.Size(50, 21);
			this.comboBoxYear.TabIndex = 0;
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
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(12, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 13);
			this.label1.TabIndex = 1;
			this.label1.Text = "Учебный год";
			// 
			// buttonImport
			// 
			this.buttonImport.Location = new System.Drawing.Point(68, 25);
			this.buttonImport.Name = "buttonImport";
			this.buttonImport.Size = new System.Drawing.Size(75, 23);
			this.buttonImport.TabIndex = 2;
			this.buttonImport.Text = "Импорт";
			this.buttonImport.UseVisualStyleBackColor = true;
			this.buttonImport.Click += new System.EventHandler(this.buttonImport_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(12, 50);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(76, 13);
			this.label2.TabIndex = 3;
			this.label2.Text = "Уже введены";
			// 
			// listBox1
			// 
			this.listBox1.DataSource = this.yearsExistBindingSource;
			this.listBox1.DisplayMember = "StudyYear";
			this.listBox1.FormattingEnabled = true;
			this.listBox1.Location = new System.Drawing.Point(12, 66);
			this.listBox1.Name = "listBox1";
			this.listBox1.Size = new System.Drawing.Size(131, 69);
			this.listBox1.TabIndex = 4;
			this.listBox1.TabStop = false;
			this.listBox1.UseTabStops = false;
			this.listBox1.ValueMember = "ID";
			// 
			// yearsQueryBindingSource
			// 
			this.yearsQueryBindingSource.DataMember = "YearsQuery";
			this.yearsQueryBindingSource.DataSource = this.mainDBDataSet1;
			// 
			// mainDBDataSet1
			// 
			this.mainDBDataSet1.DataSetName = "MainDBDataSet";
			this.mainDBDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
			// 
			// yearsExistBindingSource
			// 
			this.yearsExistBindingSource.DataMember = "YearsExistQuery";
			this.yearsExistBindingSource.DataSource = this.mainDBDataSet;
			// 
			// studyYearTableAdapter
			// 
			this.studyYearTableAdapter.ClearBeforeFill = true;
			// 
			// yearsExistTableAdapter
			// 
			this.yearsExistTableAdapter.ClearBeforeFill = true;
			// 
			// yearsQueryTableAdapter
			// 
			this.yearsQueryTableAdapter.ClearBeforeFill = true;
			// 
			// yearsExistQueryBindingSource
			// 
			this.yearsExistQueryBindingSource.DataMember = "YearsExistQuery";
			this.yearsExistQueryBindingSource.DataSource = this.mainDBDataSet1;
			// 
			// ImportForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(153, 155);
			this.Controls.Add(this.listBox1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.buttonImport);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.comboBoxYear);
			this.Name = "ImportForm";
			this.Text = "ImportForm";
			this.Load += new System.EventHandler(this.ImportForm_Load);
			((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.yearsQueryBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.yearsExistBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.yearsExistQueryBindingSource)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBoxYear;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonImport;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox listBox1;
        private MainDBDataSet mainDBDataSet;
        private System.Windows.Forms.BindingSource studyYearBindingSource;
        private MainDBDataSetTableAdapters.StudyYearTableAdapter studyYearTableAdapter;
        private System.Windows.Forms.BindingSource yearsExistBindingSource;
        private MainDBDataSetTableAdapters.YearsExistQueryTableAdapter yearsExistTableAdapter;
		private MainDBDataSet mainDBDataSet1;
		private System.Windows.Forms.BindingSource yearsQueryBindingSource;
		private MainDBDataSetTableAdapters.YearsQueryTableAdapter yearsQueryTableAdapter;
		private System.Windows.Forms.BindingSource yearsExistQueryBindingSource;
	}
}