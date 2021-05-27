namespace CafedralDB.Forms.Manage
{
    partial class ManageMainForm
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
            this.buttonSimpleDisciplines = new System.Windows.Forms.Button();
            this.buttonSpecialDisciplines = new System.Windows.Forms.Button();
            this.buttonImportSettings = new System.Windows.Forms.Button();
            this.buttonExsportSettings = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonSimpleDisciplines
            // 
            this.buttonSimpleDisciplines.Location = new System.Drawing.Point(12, 12);
            this.buttonSimpleDisciplines.Name = "buttonSimpleDisciplines";
            this.buttonSimpleDisciplines.Size = new System.Drawing.Size(75, 49);
            this.buttonSimpleDisciplines.TabIndex = 0;
            this.buttonSimpleDisciplines.Text = "Настройки обычных дисциплин";
            this.buttonSimpleDisciplines.UseVisualStyleBackColor = true;
            // 
            // buttonSpecialDisciplines
            // 
            this.buttonSpecialDisciplines.Location = new System.Drawing.Point(93, 12);
            this.buttonSpecialDisciplines.Name = "buttonSpecialDisciplines";
            this.buttonSpecialDisciplines.Size = new System.Drawing.Size(75, 49);
            this.buttonSpecialDisciplines.TabIndex = 1;
            this.buttonSpecialDisciplines.Text = "Настройки особых дисциплин";
            this.buttonSpecialDisciplines.UseVisualStyleBackColor = true;
            // 
            // buttonImportSettings
            // 
            this.buttonImportSettings.Location = new System.Drawing.Point(12, 67);
            this.buttonImportSettings.Name = "buttonImportSettings";
            this.buttonImportSettings.Size = new System.Drawing.Size(75, 53);
            this.buttonImportSettings.TabIndex = 2;
            this.buttonImportSettings.Text = "Настройки импорта";
            this.buttonImportSettings.UseVisualStyleBackColor = true;
            // 
            // buttonExsportSettings
            // 
            this.buttonExsportSettings.Location = new System.Drawing.Point(93, 67);
            this.buttonExsportSettings.Name = "buttonExsportSettings";
            this.buttonExsportSettings.Size = new System.Drawing.Size(75, 53);
            this.buttonExsportSettings.TabIndex = 3;
            this.buttonExsportSettings.Text = "Настройки экспорта";
            this.buttonExsportSettings.UseVisualStyleBackColor = true;
            // 
            // ManageMainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(174, 126);
            this.Controls.Add(this.buttonExsportSettings);
            this.Controls.Add(this.buttonImportSettings);
            this.Controls.Add(this.buttonSpecialDisciplines);
            this.Controls.Add(this.buttonSimpleDisciplines);
            this.Name = "ManageMainForm";
            this.Text = "ManageMainForm";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonSimpleDisciplines;
        private System.Windows.Forms.Button buttonSpecialDisciplines;
        private System.Windows.Forms.Button buttonImportSettings;
        private System.Windows.Forms.Button buttonExsportSettings;
    }
}