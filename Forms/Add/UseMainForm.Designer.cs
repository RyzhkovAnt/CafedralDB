namespace CafedralDB.Forms.Add
{
    partial class UseMainForm
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
			this.buttonTables = new System.Windows.Forms.Button();
			this.buttonImport = new System.Windows.Forms.Button();
			this.buttonSampleData = new System.Windows.Forms.Button();
			this.buttonAssign = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// buttonTables
			// 
			this.buttonTables.Location = new System.Drawing.Point(12, 12);
			this.buttonTables.Name = "buttonTables";
			this.buttonTables.Size = new System.Drawing.Size(94, 46);
			this.buttonTables.TabIndex = 0;
			this.buttonTables.Text = "Редактировать таблицы";
			this.buttonTables.UseVisualStyleBackColor = true;
			this.buttonTables.Click += new System.EventHandler(this.buttonTables_Click);
			// 
			// buttonImport
			// 
			this.buttonImport.Location = new System.Drawing.Point(112, 12);
			this.buttonImport.Name = "buttonImport";
			this.buttonImport.Size = new System.Drawing.Size(75, 46);
			this.buttonImport.TabIndex = 1;
			this.buttonImport.Text = "Ввести данные";
			this.buttonImport.UseVisualStyleBackColor = true;
			this.buttonImport.Click += new System.EventHandler(this.buttonImport_Click);
			// 
			// buttonSampleData
			// 
			this.buttonSampleData.Enabled = false;
			this.buttonSampleData.Location = new System.Drawing.Point(193, 12);
			this.buttonSampleData.Name = "buttonSampleData";
			this.buttonSampleData.Size = new System.Drawing.Size(75, 46);
			this.buttonSampleData.TabIndex = 2;
			this.buttonSampleData.Text = "Выборка данных";
			this.buttonSampleData.UseVisualStyleBackColor = true;
			// 
			// buttonAssign
			// 
			this.buttonAssign.Location = new System.Drawing.Point(274, 12);
			this.buttonAssign.Name = "buttonAssign";
			this.buttonAssign.Size = new System.Drawing.Size(101, 46);
			this.buttonAssign.TabIndex = 3;
			this.buttonAssign.Text = "Назначить преподавателей";
			this.buttonAssign.UseVisualStyleBackColor = true;
			this.buttonAssign.Click += new System.EventHandler(this.buttonAssign_Click);
			// 
			// UseMainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(387, 70);
			this.Controls.Add(this.buttonAssign);
			this.Controls.Add(this.buttonSampleData);
			this.Controls.Add(this.buttonImport);
			this.Controls.Add(this.buttonTables);
			this.Name = "UseMainForm";
			this.Text = "AddMainForm";
			this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonTables;
        private System.Windows.Forms.Button buttonImport;
        private System.Windows.Forms.Button buttonSampleData;
        private System.Windows.Forms.Button buttonAssign;
    }
}