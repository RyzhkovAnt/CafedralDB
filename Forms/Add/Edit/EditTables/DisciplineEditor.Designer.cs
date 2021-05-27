namespace CafedralDB.Forms.Add.Edit.EditTables
{
	partial class DisciplineEditor
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
            this.buttonSave = new System.Windows.Forms.Button();
            this.splitter2 = new System.Windows.Forms.Splitter();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.disciplineBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.mainDBDataSet = new CafedralDB.MainDBDataSet();
            this.disciplineTableAdapter = new CafedralDB.MainDBDataSetTableAdapters.DisciplineTableAdapter();
            this.iDDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.descrDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.departmentDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.disciplineTypeDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lectureCountDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.practiceCountDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.labCountDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.kRDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.kPDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.rGRDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.zachDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.ekzDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.konsDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.uchPrDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.prPrDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.predDipPrDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.konsZaochDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.gEKDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.gAKDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.gAKPredDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.dPrukDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.dopuskVKRDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.retzVKRDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.dPretzDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.aSPRukDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.mAGRukDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.mAGRetzDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.rukKafDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.nIIRDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Contr = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.disciplineBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.iDDataGridViewTextBoxColumn,
            this.descrDataGridViewTextBoxColumn,
            this.departmentDataGridViewTextBoxColumn,
            this.disciplineTypeDataGridViewTextBoxColumn,
            this.lectureCountDataGridViewTextBoxColumn,
            this.practiceCountDataGridViewTextBoxColumn,
            this.labCountDataGridViewTextBoxColumn,
            this.kRDataGridViewCheckBoxColumn,
            this.kPDataGridViewCheckBoxColumn,
            this.rGRDataGridViewCheckBoxColumn,
            this.zachDataGridViewCheckBoxColumn,
            this.ekzDataGridViewCheckBoxColumn,
            this.konsDataGridViewCheckBoxColumn,
            this.uchPrDataGridViewTextBoxColumn,
            this.prPrDataGridViewTextBoxColumn,
            this.predDipPrDataGridViewTextBoxColumn,
            this.konsZaochDataGridViewCheckBoxColumn,
            this.gEKDataGridViewCheckBoxColumn,
            this.gAKDataGridViewCheckBoxColumn,
            this.gAKPredDataGridViewCheckBoxColumn,
            this.dPrukDataGridViewCheckBoxColumn,
            this.dopuskVKRDataGridViewCheckBoxColumn,
            this.retzVKRDataGridViewCheckBoxColumn,
            this.dPretzDataGridViewCheckBoxColumn,
            this.aSPRukDataGridViewCheckBoxColumn,
            this.mAGRukDataGridViewCheckBoxColumn,
            this.mAGRetzDataGridViewCheckBoxColumn,
            this.rukKafDataGridViewCheckBoxColumn,
            this.nIIRDataGridViewTextBoxColumn,
            this.Contr});
            this.dataGridView1.DataSource = this.disciplineBindingSource;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1100, 352);
            this.dataGridView1.TabIndex = 0;
            // 
            // buttonSave
            // 
            this.buttonSave.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.buttonSave.Location = new System.Drawing.Point(507, -2);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(93, 26);
            this.buttonSave.TabIndex = 1;
            this.buttonSave.Text = "Сохранить";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
            // 
            // splitter2
            // 
            this.splitter2.Location = new System.Drawing.Point(0, 0);
            this.splitter2.Name = "splitter2";
            this.splitter2.Size = new System.Drawing.Size(3, 383);
            this.splitter2.TabIndex = 3;
            this.splitter2.TabStop = false;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(3, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.dataGridView1);
            this.splitContainer1.Panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.splitContainer1_Panel1_Paint);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.buttonSave);
            this.splitContainer1.Size = new System.Drawing.Size(1100, 383);
            this.splitContainer1.SplitterDistance = 352;
            this.splitContainer1.TabIndex = 4;
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
            // disciplineTableAdapter
            // 
            this.disciplineTableAdapter.ClearBeforeFill = true;
            // 
            // iDDataGridViewTextBoxColumn
            // 
            this.iDDataGridViewTextBoxColumn.DataPropertyName = "ID";
            this.iDDataGridViewTextBoxColumn.HeaderText = "ID";
            this.iDDataGridViewTextBoxColumn.Name = "iDDataGridViewTextBoxColumn";
            this.iDDataGridViewTextBoxColumn.Visible = false;
            // 
            // descrDataGridViewTextBoxColumn
            // 
            this.descrDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.descrDataGridViewTextBoxColumn.DataPropertyName = "Descr";
            this.descrDataGridViewTextBoxColumn.HeaderText = "Название";
            this.descrDataGridViewTextBoxColumn.Name = "descrDataGridViewTextBoxColumn";
            this.descrDataGridViewTextBoxColumn.Width = 82;
            // 
            // departmentDataGridViewTextBoxColumn
            // 
            this.departmentDataGridViewTextBoxColumn.DataPropertyName = "Department";
            this.departmentDataGridViewTextBoxColumn.HeaderText = "Каф";
            this.departmentDataGridViewTextBoxColumn.Name = "departmentDataGridViewTextBoxColumn";
            this.departmentDataGridViewTextBoxColumn.Visible = false;
            // 
            // disciplineTypeDataGridViewTextBoxColumn
            // 
            this.disciplineTypeDataGridViewTextBoxColumn.DataPropertyName = "DisciplineType";
            this.disciplineTypeDataGridViewTextBoxColumn.HeaderText = "Тип дисц.";
            this.disciplineTypeDataGridViewTextBoxColumn.Name = "disciplineTypeDataGridViewTextBoxColumn";
            this.disciplineTypeDataGridViewTextBoxColumn.Visible = false;
            // 
            // lectureCountDataGridViewTextBoxColumn
            // 
            this.lectureCountDataGridViewTextBoxColumn.DataPropertyName = "LectureCount";
            this.lectureCountDataGridViewTextBoxColumn.HeaderText = "Лек";
            this.lectureCountDataGridViewTextBoxColumn.Name = "lectureCountDataGridViewTextBoxColumn";
            this.lectureCountDataGridViewTextBoxColumn.Width = 40;
            // 
            // practiceCountDataGridViewTextBoxColumn
            // 
            this.practiceCountDataGridViewTextBoxColumn.DataPropertyName = "PracticeCount";
            this.practiceCountDataGridViewTextBoxColumn.HeaderText = "Пр";
            this.practiceCountDataGridViewTextBoxColumn.Name = "practiceCountDataGridViewTextBoxColumn";
            this.practiceCountDataGridViewTextBoxColumn.Width = 40;
            // 
            // labCountDataGridViewTextBoxColumn
            // 
            this.labCountDataGridViewTextBoxColumn.DataPropertyName = "LabCount";
            this.labCountDataGridViewTextBoxColumn.HeaderText = "Лаб";
            this.labCountDataGridViewTextBoxColumn.Name = "labCountDataGridViewTextBoxColumn";
            this.labCountDataGridViewTextBoxColumn.Width = 40;
            // 
            // kRDataGridViewCheckBoxColumn
            // 
            this.kRDataGridViewCheckBoxColumn.DataPropertyName = "KR";
            this.kRDataGridViewCheckBoxColumn.HeaderText = "КР";
            this.kRDataGridViewCheckBoxColumn.Name = "kRDataGridViewCheckBoxColumn";
            this.kRDataGridViewCheckBoxColumn.Width = 40;
            // 
            // kPDataGridViewCheckBoxColumn
            // 
            this.kPDataGridViewCheckBoxColumn.DataPropertyName = "KP";
            this.kPDataGridViewCheckBoxColumn.HeaderText = "КП";
            this.kPDataGridViewCheckBoxColumn.Name = "kPDataGridViewCheckBoxColumn";
            this.kPDataGridViewCheckBoxColumn.Width = 40;
            // 
            // rGRDataGridViewCheckBoxColumn
            // 
            this.rGRDataGridViewCheckBoxColumn.DataPropertyName = "RGR";
            this.rGRDataGridViewCheckBoxColumn.HeaderText = "РГР";
            this.rGRDataGridViewCheckBoxColumn.Name = "rGRDataGridViewCheckBoxColumn";
            this.rGRDataGridViewCheckBoxColumn.Width = 40;
            // 
            // zachDataGridViewCheckBoxColumn
            // 
            this.zachDataGridViewCheckBoxColumn.DataPropertyName = "Zach";
            this.zachDataGridViewCheckBoxColumn.HeaderText = "Зачет";
            this.zachDataGridViewCheckBoxColumn.Name = "zachDataGridViewCheckBoxColumn";
            this.zachDataGridViewCheckBoxColumn.Width = 40;
            // 
            // ekzDataGridViewCheckBoxColumn
            // 
            this.ekzDataGridViewCheckBoxColumn.DataPropertyName = "Ekz";
            this.ekzDataGridViewCheckBoxColumn.HeaderText = "Экз";
            this.ekzDataGridViewCheckBoxColumn.Name = "ekzDataGridViewCheckBoxColumn";
            this.ekzDataGridViewCheckBoxColumn.Width = 40;
            // 
            // konsDataGridViewCheckBoxColumn
            // 
            this.konsDataGridViewCheckBoxColumn.DataPropertyName = "Kons";
            this.konsDataGridViewCheckBoxColumn.HeaderText = "Конс";
            this.konsDataGridViewCheckBoxColumn.Name = "konsDataGridViewCheckBoxColumn";
            this.konsDataGridViewCheckBoxColumn.Width = 40;
            // 
            // uchPrDataGridViewTextBoxColumn
            // 
            this.uchPrDataGridViewTextBoxColumn.DataPropertyName = "UchPr";
            this.uchPrDataGridViewTextBoxColumn.HeaderText = "Уч.пр.";
            this.uchPrDataGridViewTextBoxColumn.Name = "uchPrDataGridViewTextBoxColumn";
            this.uchPrDataGridViewTextBoxColumn.Width = 40;
            // 
            // prPrDataGridViewTextBoxColumn
            // 
            this.prPrDataGridViewTextBoxColumn.DataPropertyName = "prPr";
            this.prPrDataGridViewTextBoxColumn.HeaderText = "Произв. пр.";
            this.prPrDataGridViewTextBoxColumn.Name = "prPrDataGridViewTextBoxColumn";
            this.prPrDataGridViewTextBoxColumn.Width = 40;
            // 
            // predDipPrDataGridViewTextBoxColumn
            // 
            this.predDipPrDataGridViewTextBoxColumn.DataPropertyName = "predDipPr";
            this.predDipPrDataGridViewTextBoxColumn.HeaderText = "Преддип. пр.";
            this.predDipPrDataGridViewTextBoxColumn.Name = "predDipPrDataGridViewTextBoxColumn";
            this.predDipPrDataGridViewTextBoxColumn.Width = 40;
            // 
            // konsZaochDataGridViewCheckBoxColumn
            // 
            this.konsZaochDataGridViewCheckBoxColumn.DataPropertyName = "KonsZaoch";
            this.konsZaochDataGridViewCheckBoxColumn.HeaderText = "Конс. заоч.";
            this.konsZaochDataGridViewCheckBoxColumn.Name = "konsZaochDataGridViewCheckBoxColumn";
            this.konsZaochDataGridViewCheckBoxColumn.Width = 40;
            // 
            // gEKDataGridViewCheckBoxColumn
            // 
            this.gEKDataGridViewCheckBoxColumn.DataPropertyName = "GEK";
            this.gEKDataGridViewCheckBoxColumn.HeaderText = "ГЭК";
            this.gEKDataGridViewCheckBoxColumn.Name = "gEKDataGridViewCheckBoxColumn";
            this.gEKDataGridViewCheckBoxColumn.Width = 40;
            // 
            // gAKDataGridViewCheckBoxColumn
            // 
            this.gAKDataGridViewCheckBoxColumn.DataPropertyName = "GAK";
            this.gAKDataGridViewCheckBoxColumn.HeaderText = "ГАК";
            this.gAKDataGridViewCheckBoxColumn.Name = "gAKDataGridViewCheckBoxColumn";
            this.gAKDataGridViewCheckBoxColumn.Width = 40;
            // 
            // gAKPredDataGridViewCheckBoxColumn
            // 
            this.gAKPredDataGridViewCheckBoxColumn.DataPropertyName = "GAKPred";
            this.gAKPredDataGridViewCheckBoxColumn.HeaderText = "Гак предс";
            this.gAKPredDataGridViewCheckBoxColumn.Name = "gAKPredDataGridViewCheckBoxColumn";
            this.gAKPredDataGridViewCheckBoxColumn.Width = 40;
            // 
            // dPrukDataGridViewCheckBoxColumn
            // 
            this.dPrukDataGridViewCheckBoxColumn.DataPropertyName = "DPruk";
            this.dPrukDataGridViewCheckBoxColumn.HeaderText = "ДП рук";
            this.dPrukDataGridViewCheckBoxColumn.Name = "dPrukDataGridViewCheckBoxColumn";
            this.dPrukDataGridViewCheckBoxColumn.Width = 40;
            // 
            // dopuskVKRDataGridViewCheckBoxColumn
            // 
            this.dopuskVKRDataGridViewCheckBoxColumn.DataPropertyName = "DopuskVKR";
            this.dopuskVKRDataGridViewCheckBoxColumn.HeaderText = "Допуск ВКР";
            this.dopuskVKRDataGridViewCheckBoxColumn.Name = "dopuskVKRDataGridViewCheckBoxColumn";
            this.dopuskVKRDataGridViewCheckBoxColumn.Width = 40;
            // 
            // retzVKRDataGridViewCheckBoxColumn
            // 
            this.retzVKRDataGridViewCheckBoxColumn.DataPropertyName = "RetzVKR";
            this.retzVKRDataGridViewCheckBoxColumn.HeaderText = "Рец ВКР";
            this.retzVKRDataGridViewCheckBoxColumn.Name = "retzVKRDataGridViewCheckBoxColumn";
            this.retzVKRDataGridViewCheckBoxColumn.Width = 40;
            // 
            // dPretzDataGridViewCheckBoxColumn
            // 
            this.dPretzDataGridViewCheckBoxColumn.DataPropertyName = "DPretz";
            this.dPretzDataGridViewCheckBoxColumn.HeaderText = "ДП рец.";
            this.dPretzDataGridViewCheckBoxColumn.Name = "dPretzDataGridViewCheckBoxColumn";
            this.dPretzDataGridViewCheckBoxColumn.Width = 40;
            // 
            // aSPRukDataGridViewCheckBoxColumn
            // 
            this.aSPRukDataGridViewCheckBoxColumn.DataPropertyName = "ASPRuk";
            this.aSPRukDataGridViewCheckBoxColumn.HeaderText = "АСП рук.";
            this.aSPRukDataGridViewCheckBoxColumn.Name = "aSPRukDataGridViewCheckBoxColumn";
            this.aSPRukDataGridViewCheckBoxColumn.Width = 40;
            // 
            // mAGRukDataGridViewCheckBoxColumn
            // 
            this.mAGRukDataGridViewCheckBoxColumn.DataPropertyName = "MAGRuk";
            this.mAGRukDataGridViewCheckBoxColumn.HeaderText = "МАГ рук";
            this.mAGRukDataGridViewCheckBoxColumn.Name = "mAGRukDataGridViewCheckBoxColumn";
            this.mAGRukDataGridViewCheckBoxColumn.Width = 40;
            // 
            // mAGRetzDataGridViewCheckBoxColumn
            // 
            this.mAGRetzDataGridViewCheckBoxColumn.DataPropertyName = "MAGRetz";
            this.mAGRetzDataGridViewCheckBoxColumn.HeaderText = "МАГ рец";
            this.mAGRetzDataGridViewCheckBoxColumn.Name = "mAGRetzDataGridViewCheckBoxColumn";
            this.mAGRetzDataGridViewCheckBoxColumn.Width = 40;
            // 
            // rukKafDataGridViewCheckBoxColumn
            // 
            this.rukKafDataGridViewCheckBoxColumn.DataPropertyName = "RukKaf";
            this.rukKafDataGridViewCheckBoxColumn.HeaderText = "Рук. каф.";
            this.rukKafDataGridViewCheckBoxColumn.Name = "rukKafDataGridViewCheckBoxColumn";
            this.rukKafDataGridViewCheckBoxColumn.Width = 40;
            // 
            // nIIRDataGridViewTextBoxColumn
            // 
            this.nIIRDataGridViewTextBoxColumn.DataPropertyName = "NIIR";
            this.nIIRDataGridViewTextBoxColumn.HeaderText = "НИИР";
            this.nIIRDataGridViewTextBoxColumn.Name = "nIIRDataGridViewTextBoxColumn";
            this.nIIRDataGridViewTextBoxColumn.Width = 40;
            // 
            // Contr
            // 
            this.Contr.DataPropertyName = "Contr";
            this.Contr.HeaderText = "Договор";
            this.Contr.Name = "Contr";
            this.Contr.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // DisciplineEditor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1103, 383);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.splitter2);
            this.Name = "DisciplineEditor";
            this.Text = "DisciplineEditor";
            this.Load += new System.EventHandler(this.DisciplineEditor_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.disciplineBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.mainDBDataSet)).EndInit();
            this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.DataGridView dataGridView1;
		private MainDBDataSet mainDBDataSet;
		private System.Windows.Forms.BindingSource disciplineBindingSource;
		private MainDBDataSetTableAdapters.DisciplineTableAdapter disciplineTableAdapter;
		private System.Windows.Forms.Button buttonSave;
		private System.Windows.Forms.Splitter splitter2;
		private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.DataGridViewTextBoxColumn iDDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn descrDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn departmentDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn disciplineTypeDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn lectureCountDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn practiceCountDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn labCountDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn kRDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn kPDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn rGRDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn zachDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ekzDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn konsDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn uchPrDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn prPrDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn predDipPrDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn konsZaochDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn gEKDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn gAKDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn gAKPredDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn dPrukDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn dopuskVKRDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn retzVKRDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn dPretzDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn aSPRukDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn mAGRukDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn mAGRetzDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn rukKafDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn nIIRDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Contr;
    }
}