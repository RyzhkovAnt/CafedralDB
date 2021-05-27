namespace CafedralDB.Forms.Export
{
    partial class ExportMainForm
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
            this.buttonSemester = new System.Windows.Forms.Button();
            this.buttonIndPlan = new System.Windows.Forms.Button();
            this.buttonWorkload = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonSemester
            // 
            this.buttonSemester.Location = new System.Drawing.Point(12, 12);
            this.buttonSemester.Name = "buttonSemester";
            this.buttonSemester.Size = new System.Drawing.Size(59, 34);
            this.buttonSemester.TabIndex = 0;
            this.buttonSemester.Text = "Отчет за семестр";
            this.buttonSemester.UseVisualStyleBackColor = true;
            this.buttonSemester.Click += new System.EventHandler(this.buttonSemester_Click);
            // 
            // buttonIndPlan
            // 
            this.buttonIndPlan.Location = new System.Drawing.Point(77, 12);
            this.buttonIndPlan.Name = "buttonIndPlan";
            this.buttonIndPlan.Size = new System.Drawing.Size(131, 34);
            this.buttonIndPlan.TabIndex = 1;
            this.buttonIndPlan.Text = "Индивидуальный план";
            this.buttonIndPlan.UseVisualStyleBackColor = true;
            this.buttonIndPlan.Click += new System.EventHandler(this.buttonIndPlan_Click);
            // 
            // buttonWorkload
            // 
            this.buttonWorkload.Location = new System.Drawing.Point(214, 12);
            this.buttonWorkload.Name = "buttonWorkload";
            this.buttonWorkload.Size = new System.Drawing.Size(75, 34);
            this.buttonWorkload.TabIndex = 2;
            this.buttonWorkload.Text = "Нагрузка";
            this.buttonWorkload.UseVisualStyleBackColor = true;
            this.buttonWorkload.Click += new System.EventHandler(this.buttonWorkload_Click);
            // 
            // ExportMainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(300, 58);
            this.Controls.Add(this.buttonWorkload);
            this.Controls.Add(this.buttonIndPlan);
            this.Controls.Add(this.buttonSemester);
            this.Name = "ExportMainForm";
            this.Text = "ExportMainForm";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonSemester;
        private System.Windows.Forms.Button buttonIndPlan;
        private System.Windows.Forms.Button buttonWorkload;
    }
}