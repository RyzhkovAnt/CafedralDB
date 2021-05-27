namespace CafedralDB.Forms.Export
{
    partial class ExportWorkload
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
			this.label1 = new System.Windows.Forms.Label();
			this.comboBoxYear = new System.Windows.Forms.ComboBox();
			this.studyYearBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.mainDBDataSet = new CafedralDB.MainDBDataSet();
			this.buttonExport = new System.Windows.Forms.Button();
			this.studyYearTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.StudyYearTableAdapter();
			((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(97, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 13);
			this.label1.TabIndex = 6;
			this.label1.Text = "Учебный год";
			// 
			// comboBoxYear
			// 
			this.comboBoxYear.DataSource = this.studyYearBindingSource;
			this.comboBoxYear.DisplayMember = "StudyYear";
			this.comboBoxYear.FormattingEnabled = true;
			this.comboBoxYear.Location = new System.Drawing.Point(82, 25);
			this.comboBoxYear.Name = "comboBoxYear";
			this.comboBoxYear.Size = new System.Drawing.Size(121, 21);
			this.comboBoxYear.TabIndex = 5;
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
			// buttonExport
			// 
			this.buttonExport.Location = new System.Drawing.Point(100, 50);
			this.buttonExport.Name = "buttonExport";
			this.buttonExport.Size = new System.Drawing.Size(75, 23);
			this.buttonExport.TabIndex = 4;
			this.buttonExport.Text = "Экспорт";
			this.buttonExport.UseVisualStyleBackColor = true;
			this.buttonExport.Click += new System.EventHandler(this.buttonExport_Click);
			// 
			// studyYearTableAdapter
			// 
			this.studyYearTableAdapter.ClearBeforeFill = true;
			// 
			// ExportWorkload
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(284, 81);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.comboBoxYear);
			this.Controls.Add(this.buttonExport);
			this.Name = "ExportWorkload";
			this.Text = "ExportWorkload";
			this.Load += new System.EventHandler(this.ExportWorkload_Load);
			((System.ComponentModel.ISupportInitialize)(this.studyYearBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBoxYear;
        private System.Windows.Forms.Button buttonExport;
        private MainDBDataSet mainDBDataSet;
        private System.Windows.Forms.BindingSource studyYearBindingSource;
        private MainDBDataSetTableAdapters.StudyYearTableAdapter studyYearTableAdapter;
    }
}