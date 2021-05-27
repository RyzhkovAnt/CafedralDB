namespace CafedralDB.Forms.Add.Edit
{
	partial class EditorSelector
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
			this.comboBoxTableSelection = new System.Windows.Forms.ComboBox();
			this.button1 = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// comboBoxTableSelection
			// 
			this.comboBoxTableSelection.FormattingEnabled = true;
			this.comboBoxTableSelection.Items.AddRange(new object[] {
            "Ученая степень",
            "Ученое звание",
            "Кафедра",
            "Дисциплина",
            "Тип дисциплины",
            "Сотрудник",
            "Факультет",
            "Группа",
            "Квалификация",
            "Семестр",
            "Специальность",
            "Студенты",
            "Форма обучения",
            "Учебный год",
            "Должность сотрудника",
            "Нагрузка",
            "Назначение нагрузки"});
			this.comboBoxTableSelection.Location = new System.Drawing.Point(12, 12);
			this.comboBoxTableSelection.Name = "comboBoxTableSelection";
			this.comboBoxTableSelection.Size = new System.Drawing.Size(121, 21);
			this.comboBoxTableSelection.TabIndex = 0;
			this.comboBoxTableSelection.Text = "Ученая степень";
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(139, 12);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(99, 23);
			this.button1.TabIndex = 1;
			this.button1.Text = "Редактировать";
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// EditorSelector
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(250, 41);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.comboBoxTableSelection);
			this.Name = "EditorSelector";
			this.Text = "EditorSelector";
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.ComboBox comboBoxTableSelection;
		private System.Windows.Forms.Button button1;
	}
}