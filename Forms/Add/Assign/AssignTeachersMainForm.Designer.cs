namespace CafedralDB.Forms.Add.Assign
{
    partial class AssignTeachersMainForm
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
            this.buttonSimpleAssign = new System.Windows.Forms.Button();
            this.buttonSpecialAssign = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonSimpleAssign
            // 
            this.buttonSimpleAssign.Location = new System.Drawing.Point(12, 12);
            this.buttonSimpleAssign.Name = "buttonSimpleAssign";
            this.buttonSimpleAssign.Size = new System.Drawing.Size(80, 36);
            this.buttonSimpleAssign.TabIndex = 0;
            this.buttonSimpleAssign.Text = "Простые дисциплины";
            this.buttonSimpleAssign.UseVisualStyleBackColor = true;
            this.buttonSimpleAssign.Click += new System.EventHandler(this.buttonSimpleAssign_Click);
            // 
            // buttonSpecialAssign
            // 
            this.buttonSpecialAssign.Location = new System.Drawing.Point(98, 12);
            this.buttonSpecialAssign.Name = "buttonSpecialAssign";
            this.buttonSpecialAssign.Size = new System.Drawing.Size(81, 36);
            this.buttonSpecialAssign.TabIndex = 1;
            this.buttonSpecialAssign.Text = "Особые дисциплины";
            this.buttonSpecialAssign.UseVisualStyleBackColor = true;
            this.buttonSpecialAssign.Click += new System.EventHandler(this.buttonSpecialAssign_Click);
            // 
            // AssignTeachersMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(188, 55);
            this.Controls.Add(this.buttonSpecialAssign);
            this.Controls.Add(this.buttonSimpleAssign);
            this.Name = "AssignTeachersMain";
            this.Text = "AssignTeachersMain";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonSimpleAssign;
        private System.Windows.Forms.Button buttonSpecialAssign;
    }
}