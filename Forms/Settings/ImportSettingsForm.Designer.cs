namespace CafedralDB.Forms.Settings
{
	partial class ImportSettingsForm
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
            this.buttonSave = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxStartReadingRow = new System.Windows.Forms.TextBox();
            this.textBoxGroupColumn = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxSemesterColumn = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxWeekColumn = new System.Windows.Forms.TextBox();
            this.label = new System.Windows.Forms.Label();
            this.textBoxDisciplineNameColumn = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textBoxLecturesColumn = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.textBoxLabsColumn = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.textBoxPracticesColumn = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.textBoxKzColumn = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.textBoxKrColumn = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.textBoxKpColumn = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.textBoxEkzColumn = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.textBoxZachColumn = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.textBoxOtherColumn = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBoxContractColumn = new System.Windows.Forms.TextBox();
            this.toDefaultButton = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonSave
            // 
            this.buttonSave.Location = new System.Drawing.Point(228, 251);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(75, 23);
            this.buttonSave.TabIndex = 0;
            this.buttonSave.Text = "Сохранить";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(98, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Стартовая строка";
            // 
            // textBoxStartReadingRow
            // 
            this.textBoxStartReadingRow.Location = new System.Drawing.Point(114, 6);
            this.textBoxStartReadingRow.Name = "textBoxStartReadingRow";
            this.textBoxStartReadingRow.Size = new System.Drawing.Size(38, 20);
            this.textBoxStartReadingRow.TabIndex = 2;
            this.textBoxStartReadingRow.TextChanged += new System.EventHandler(this.textBoxStartReadingRow_TextChanged);
            // 
            // textBoxGroupColumn
            // 
            this.textBoxGroupColumn.Location = new System.Drawing.Point(107, 23);
            this.textBoxGroupColumn.Name = "textBoxGroupColumn";
            this.textBoxGroupColumn.Size = new System.Drawing.Size(38, 20);
            this.textBoxGroupColumn.TabIndex = 4;
            this.textBoxGroupColumn.TextChanged += new System.EventHandler(this.textBoxGroupColumn_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 26);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(42, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Группа";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(5, 52);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(51, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Семестр";
            // 
            // textBoxSemesterColumn
            // 
            this.textBoxSemesterColumn.Location = new System.Drawing.Point(107, 49);
            this.textBoxSemesterColumn.Name = "textBoxSemesterColumn";
            this.textBoxSemesterColumn.Size = new System.Drawing.Size(38, 20);
            this.textBoxSemesterColumn.TabIndex = 4;
            this.textBoxSemesterColumn.TextChanged += new System.EventHandler(this.textBoxSemesterColumn_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(5, 78);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(45, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "Недели";
            // 
            // textBoxWeekColumn
            // 
            this.textBoxWeekColumn.Location = new System.Drawing.Point(107, 75);
            this.textBoxWeekColumn.Name = "textBoxWeekColumn";
            this.textBoxWeekColumn.Size = new System.Drawing.Size(38, 20);
            this.textBoxWeekColumn.TabIndex = 4;
            this.textBoxWeekColumn.TextChanged += new System.EventHandler(this.textBoxWeekColumn_TextChanged);
            // 
            // label
            // 
            this.label.AutoSize = true;
            this.label.Location = new System.Drawing.Point(5, 104);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(70, 13);
            this.label.TabIndex = 3;
            this.label.Text = "Дисциплина";
            // 
            // textBoxDisciplineNameColumn
            // 
            this.textBoxDisciplineNameColumn.Location = new System.Drawing.Point(107, 101);
            this.textBoxDisciplineNameColumn.Name = "textBoxDisciplineNameColumn";
            this.textBoxDisciplineNameColumn.Size = new System.Drawing.Size(38, 20);
            this.textBoxDisciplineNameColumn.TabIndex = 4;
            this.textBoxDisciplineNameColumn.TextChanged += new System.EventHandler(this.textBoxDisciplineNameColumn_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(5, 130);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(45, 13);
            this.label6.TabIndex = 3;
            this.label6.Text = "Лекции";
            // 
            // textBoxLecturesColumn
            // 
            this.textBoxLecturesColumn.Location = new System.Drawing.Point(107, 127);
            this.textBoxLecturesColumn.Name = "textBoxLecturesColumn";
            this.textBoxLecturesColumn.Size = new System.Drawing.Size(38, 20);
            this.textBoxLecturesColumn.TabIndex = 4;
            this.textBoxLecturesColumn.TextChanged += new System.EventHandler(this.textBoxLecturesColumn_TextChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(5, 156);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(71, 13);
            this.label7.TabIndex = 3;
            this.label7.Text = "Лабораторн.";
            // 
            // textBoxLabsColumn
            // 
            this.textBoxLabsColumn.Location = new System.Drawing.Point(107, 153);
            this.textBoxLabsColumn.Name = "textBoxLabsColumn";
            this.textBoxLabsColumn.Size = new System.Drawing.Size(38, 20);
            this.textBoxLabsColumn.TabIndex = 4;
            this.textBoxLabsColumn.TextChanged += new System.EventHandler(this.textBoxLabsColumn_TextChanged);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(162, 27);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(56, 13);
            this.label8.TabIndex = 1;
            this.label8.Text = "Практики";
            // 
            // textBoxPracticesColumn
            // 
            this.textBoxPracticesColumn.Location = new System.Drawing.Point(253, 24);
            this.textBoxPracticesColumn.Name = "textBoxPracticesColumn";
            this.textBoxPracticesColumn.Size = new System.Drawing.Size(38, 20);
            this.textBoxPracticesColumn.TabIndex = 2;
            this.textBoxPracticesColumn.TextChanged += new System.EventHandler(this.textBoxPracticesColumn_TextChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(161, 53);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(88, 13);
            this.label9.TabIndex = 3;
            this.label9.Text = "Контрольн. зад.";
            // 
            // textBoxKzColumn
            // 
            this.textBoxKzColumn.Location = new System.Drawing.Point(253, 50);
            this.textBoxKzColumn.Name = "textBoxKzColumn";
            this.textBoxKzColumn.Size = new System.Drawing.Size(38, 20);
            this.textBoxKzColumn.TabIndex = 4;
            this.textBoxKzColumn.TextChanged += new System.EventHandler(this.textBoxKzColumn_TextChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(162, 79);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(21, 13);
            this.label10.TabIndex = 3;
            this.label10.Text = "КР";
            // 
            // textBoxKrColumn
            // 
            this.textBoxKrColumn.Location = new System.Drawing.Point(253, 76);
            this.textBoxKrColumn.Name = "textBoxKrColumn";
            this.textBoxKrColumn.Size = new System.Drawing.Size(38, 20);
            this.textBoxKrColumn.TabIndex = 4;
            this.textBoxKrColumn.TextChanged += new System.EventHandler(this.textBoxKrColumn_TextChanged);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(161, 105);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(22, 13);
            this.label11.TabIndex = 3;
            this.label11.Text = "КП";
            // 
            // textBoxKpColumn
            // 
            this.textBoxKpColumn.Location = new System.Drawing.Point(253, 102);
            this.textBoxKpColumn.Name = "textBoxKpColumn";
            this.textBoxKpColumn.Size = new System.Drawing.Size(38, 20);
            this.textBoxKpColumn.TabIndex = 4;
            this.textBoxKpColumn.TextChanged += new System.EventHandler(this.textBoxKpColumn_TextChanged);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(161, 132);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(52, 13);
            this.label12.TabIndex = 3;
            this.label12.Text = "Экзамен";
            // 
            // textBoxEkzColumn
            // 
            this.textBoxEkzColumn.Location = new System.Drawing.Point(253, 128);
            this.textBoxEkzColumn.Name = "textBoxEkzColumn";
            this.textBoxEkzColumn.Size = new System.Drawing.Size(38, 20);
            this.textBoxEkzColumn.TabIndex = 4;
            this.textBoxEkzColumn.TextChanged += new System.EventHandler(this.textBoxEkzColumn_TextChanged);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(161, 156);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(36, 13);
            this.label13.TabIndex = 3;
            this.label13.Text = "Зачет";
            // 
            // textBoxZachColumn
            // 
            this.textBoxZachColumn.Location = new System.Drawing.Point(253, 154);
            this.textBoxZachColumn.Name = "textBoxZachColumn";
            this.textBoxZachColumn.Size = new System.Drawing.Size(38, 20);
            this.textBoxZachColumn.TabIndex = 4;
            this.textBoxZachColumn.TextChanged += new System.EventHandler(this.textBoxZachColumn_TextChanged);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(161, 180);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(100, 13);
            this.label14.TabIndex = 3;
            this.label14.Text = "Другое(depricated)";
            // 
            // textBoxOtherColumn
            // 
            this.textBoxOtherColumn.Location = new System.Drawing.Point(253, 180);
            this.textBoxOtherColumn.Name = "textBoxOtherColumn";
            this.textBoxOtherColumn.Size = new System.Drawing.Size(38, 20);
            this.textBoxOtherColumn.TabIndex = 4;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.textBoxContractColumn);
            this.groupBox1.Controls.Add(this.textBoxLabsColumn);
            this.groupBox1.Controls.Add(this.textBoxOtherColumn);
            this.groupBox1.Controls.Add(this.textBoxZachColumn);
            this.groupBox1.Controls.Add(this.label14);
            this.groupBox1.Controls.Add(this.textBoxEkzColumn);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.textBoxKpColumn);
            this.groupBox1.Controls.Add(this.textBoxLecturesColumn);
            this.groupBox1.Controls.Add(this.textBoxKrColumn);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.textBoxKzColumn);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.textBoxPracticesColumn);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.textBoxDisciplineNameColumn);
            this.groupBox1.Controls.Add(this.label);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.textBoxWeekColumn);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.textBoxSemesterColumn);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.textBoxGroupColumn);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Location = new System.Drawing.Point(12, 32);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(298, 213);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Номера столбцов";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 182);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(107, 13);
            this.label5.TabIndex = 6;
            this.label5.Text = "Договор(depricated)";
            // 
            // textBoxContractColumn
            // 
            this.textBoxContractColumn.Location = new System.Drawing.Point(108, 179);
            this.textBoxContractColumn.Name = "textBoxContractColumn";
            this.textBoxContractColumn.Size = new System.Drawing.Size(37, 20);
            this.textBoxContractColumn.TabIndex = 5;
            this.textBoxContractColumn.TextChanged += new System.EventHandler(this.textBoxContractColumn_TextChanged);
            // 
            // toDefaultButton
            // 
            this.toDefaultButton.Location = new System.Drawing.Point(15, 251);
            this.toDefaultButton.Name = "toDefaultButton";
            this.toDefaultButton.Size = new System.Drawing.Size(110, 23);
            this.toDefaultButton.TabIndex = 6;
            this.toDefaultButton.Text = "по умолчанию";
            this.toDefaultButton.UseVisualStyleBackColor = true;
            this.toDefaultButton.Click += new System.EventHandler(this.toDefaultButton_Click);
            // 
            // ImportSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(317, 279);
            this.Controls.Add(this.toDefaultButton);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.textBoxStartReadingRow);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonSave);
            this.Name = "ImportSettingsForm";
            this.Text = "ImportSettingsForm";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button buttonSave;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox textBoxStartReadingRow;
		private System.Windows.Forms.TextBox textBoxGroupColumn;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox textBoxSemesterColumn;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox textBoxWeekColumn;
		private System.Windows.Forms.Label label;
		private System.Windows.Forms.TextBox textBoxDisciplineNameColumn;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox textBoxLecturesColumn;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.TextBox textBoxLabsColumn;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.TextBox textBoxPracticesColumn;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.TextBox textBoxKzColumn;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.TextBox textBoxKrColumn;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.TextBox textBoxKpColumn;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.TextBox textBoxEkzColumn;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.TextBox textBoxZachColumn;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.TextBox textBoxOtherColumn;
		private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBoxContractColumn;
        private System.Windows.Forms.Button toDefaultButton;
    }
}