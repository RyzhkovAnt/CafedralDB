using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CafedralDB.Forms.Export
{
    public partial class ExportMainForm : Form
    {
        public ExportMainForm()
        {
            InitializeComponent();
        }

        private void buttonSemester_Click(object sender, EventArgs e)
        {
            ExportSemester semesterForm = new ExportSemester();
            semesterForm.Show();
        }

        private void buttonIndPlan_Click(object sender, EventArgs e)
        {
            ExportIndPlan indPlanForm = new ExportIndPlan();
            indPlanForm.Show();
        }

        private void buttonWorkload_Click(object sender, EventArgs e)
        {
            ExportWorkload workloadForm = new ExportWorkload();
            workloadForm.Show();
        }
    }
}
