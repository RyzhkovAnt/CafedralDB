namespace CafedralDB.Forms.Add.Import
{
    partial class CheckImportForm
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.buttonYes = new System.Windows.Forms.Button();
            this.buttonNo = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SpecialityColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SemesterColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.WeekColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DisciplineColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LecturesColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PracticeColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LabsColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.KzColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.KrColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.KpColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.EkzColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.ZachColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.SpecialityColumn,
            this.SemesterColumn,
            this.WeekColumn,
            this.DisciplineColumn,
            this.LecturesColumn,
            this.PracticeColumn,
            this.LabsColumn,
            this.KzColumn,
            this.KrColumn,
            this.KpColumn,
            this.EkzColumn,
            this.ZachColumn});
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(911, 284);
            this.dataGridView1.TabIndex = 1;
            // 
            // buttonYes
            // 
            this.buttonYes.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonYes.Location = new System.Drawing.Point(184, 302);
            this.buttonYes.Name = "buttonYes";
            this.buttonYes.Size = new System.Drawing.Size(124, 61);
            this.buttonYes.TabIndex = 2;
            this.buttonYes.Text = "Да";
            this.buttonYes.UseVisualStyleBackColor = true;
            // 
            // buttonNo
            // 
            this.buttonNo.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonNo.Location = new System.Drawing.Point(600, 302);
            this.buttonNo.Name = "buttonNo";
            this.buttonNo.Size = new System.Drawing.Size(102, 62);
            this.buttonNo.TabIndex = 3;
            this.buttonNo.Text = "Нет";
            this.buttonNo.UseVisualStyleBackColor = true;
            this.buttonNo.Click += new System.EventHandler(this.buttonNo_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(314, 302);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(280, 55);
            this.label1.TabIndex = 4;
            this.label1.Text = "Все верно?";
            // 
            // SpecialityColumn
            // 
            this.SpecialityColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.SpecialityColumn.HeaderText = "Спец-сть";
            this.SpecialityColumn.MinimumWidth = 20;
            this.SpecialityColumn.Name = "SpecialityColumn";
            this.SpecialityColumn.ReadOnly = true;
            this.SpecialityColumn.Width = 60;
            // 
            // SemesterColumn
            // 
            this.SemesterColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.SemesterColumn.HeaderText = "Семестр";
            this.SemesterColumn.Name = "SemesterColumn";
            this.SemesterColumn.ReadOnly = true;
            this.SemesterColumn.Width = 55;
            // 
            // WeekColumn
            // 
            this.WeekColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.WeekColumn.HeaderText = "Недель";
            this.WeekColumn.Name = "WeekColumn";
            this.WeekColumn.ReadOnly = true;
            this.WeekColumn.Width = 50;
            // 
            // DisciplineColumn
            // 
            this.DisciplineColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.DisciplineColumn.HeaderText = "Наименование дисциплины";
            this.DisciplineColumn.Name = "DisciplineColumn";
            this.DisciplineColumn.ReadOnly = true;
            // 
            // LecturesColumn
            // 
            this.LecturesColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.LecturesColumn.HeaderText = "Лекций";
            this.LecturesColumn.Name = "LecturesColumn";
            this.LecturesColumn.ReadOnly = true;
            this.LecturesColumn.Width = 50;
            // 
            // PracticeColumn
            // 
            this.PracticeColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.PracticeColumn.HeaderText = "Практик";
            this.PracticeColumn.Name = "PracticeColumn";
            this.PracticeColumn.ReadOnly = true;
            this.PracticeColumn.Width = 53;
            // 
            // LabsColumn
            // 
            this.LabsColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.LabsColumn.HeaderText = "Лаб";
            this.LabsColumn.Name = "LabsColumn";
            this.LabsColumn.ReadOnly = true;
            this.LabsColumn.Width = 36;
            // 
            // KzColumn
            // 
            this.KzColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.KzColumn.HeaderText = "КЗ";
            this.KzColumn.Name = "KzColumn";
            this.KzColumn.ReadOnly = true;
            this.KzColumn.Width = 27;
            // 
            // KrColumn
            // 
            this.KrColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.KrColumn.HeaderText = "КР";
            this.KrColumn.Name = "KrColumn";
            this.KrColumn.ReadOnly = true;
            this.KrColumn.Width = 27;
            // 
            // KpColumn
            // 
            this.KpColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.KpColumn.HeaderText = "КП";
            this.KpColumn.Name = "KpColumn";
            this.KpColumn.ReadOnly = true;
            this.KpColumn.Width = 28;
            // 
            // EkzColumn
            // 
            this.EkzColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.EkzColumn.HeaderText = "Экз";
            this.EkzColumn.Name = "EkzColumn";
            this.EkzColumn.ReadOnly = true;
            this.EkzColumn.Width = 32;
            // 
            // ZachColumn
            // 
            this.ZachColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.ZachColumn.HeaderText = "Зачет";
            this.ZachColumn.Name = "ZachColumn";
            this.ZachColumn.ReadOnly = true;
            this.ZachColumn.Width = 42;
            // 
            // CheckImportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(935, 376);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonNo);
            this.Controls.Add(this.buttonYes);
            this.Controls.Add(this.dataGridView1);
            this.Name = "CheckImportForm";
            this.Text = "CheckImportForm";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button buttonYes;
        private System.Windows.Forms.Button buttonNo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridViewTextBoxColumn SpecialityColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn SemesterColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn WeekColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn DisciplineColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn LecturesColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn PracticeColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn LabsColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn KzColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn KrColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn KpColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn EkzColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ZachColumn;
    }
}