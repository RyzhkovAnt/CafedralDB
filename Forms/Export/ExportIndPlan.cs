using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CafedralDB.SourceCode.Model.Exporter;

namespace CafedralDB.Forms.Export
{
    public partial class ExportIndPlan : Form
    {
        public ExportIndPlan()
        {
            InitializeComponent();
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            string filePath;
            string year = comboBoxYear.SelectedValue.ToString();
			int teacherID = Convert.ToInt32(comboBoxTeacher.SelectedValue);


            SourceCode.Model.Exporter.ExportIndPlan.Export(teacherID, year);
            //IndPlan.ExportIndPlan(teacherID, year);
        }

        private void ExportIndPlan_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'mainDBDataSet.Employee' table. You can move, or remove it, as needed.
            this.employeeTableAdapter.Fill(this.mainDBDataSet.Employee);
            // TODO: This line of code loads data into the 'mainDBDataSet.StudyYear' table. You can move, or remove it, as needed.
            this.studyYearTableAdapter.Fill(this.mainDBDataSet.StudyYear);

        }
    }
}
