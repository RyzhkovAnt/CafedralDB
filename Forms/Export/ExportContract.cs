using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CafedralDB.SourceCode;

namespace CafedralDB.Forms.Export
{
    public partial class ExportContract : Form
    {
        public ExportContract()
        {
            InitializeComponent();
            Semester_comboBox.DataSource=Enum.GetValues(typeof(Utilities.SemesterName));
            Semester_comboBox.SelectedIndex=0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var selectedSemester = (Utilities.SemesterName)Semester_comboBox.SelectedItem;
            SourceCode.Model.Exporter.ExportContract.Export(Convert.ToInt32(Teacher_comboBox.SelectedValue), Year_comboBox.SelectedValue.ToString(), selectedSemester);
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
