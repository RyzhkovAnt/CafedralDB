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
    public partial class ExportContract : Form
    {
        public ExportContract()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Contract.ExportContract(Convert.ToInt32(Teacher_comboBox.SelectedValue),
                        Year_comboBox.SelectedValue.ToString(),Semester_comboBox.SelectedItem.ToString());
        }

        private void ExportContract_Load(object sender, EventArgs e)
        {
            this.employeeTableAdapter1.Fill(this.mainDBDataSet1.Employee);
            this.studyYearTableAdapter1.Fill(this.mainDBDataSet1.StudyYear);
            Semester_comboBox.SelectedIndex = 0;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
