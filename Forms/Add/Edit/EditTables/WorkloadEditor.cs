using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CafedralDB.Forms.Add.Edit.EditTables
{
	public partial class WorkloadEditor : Form
	{
		public WorkloadEditor()
		{
			InitializeComponent();
		}

		private void WorkloadEditor_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.Group' table. You can move, or remove it, as needed.
			this.groupTableAdapter.Fill(this.mainDBDataSet.Group);
			// TODO: This line of code loads data into the 'mainDBDataSet.Semester' table. You can move, or remove it, as needed.
			this.semesterTableAdapter.Fill(this.mainDBDataSet.Semester);
			// TODO: This line of code loads data into the 'mainDBDataSet.StudyYear' table. You can move, or remove it, as needed.
			this.studyYearTableAdapter.Fill(this.mainDBDataSet.StudyYear);
			// TODO: This line of code loads data into the 'mainDBDataSet.Discipline' table. You can move, or remove it, as needed.
			this.disciplineTableAdapter.Fill(this.mainDBDataSet.Discipline);
			// TODO: This line of code loads data into the 'mainDBDataSet.Workload' table. You can move, or remove it, as needed.
			this.workloadTableAdapter.Fill(this.mainDBDataSet.Workload);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			this.workloadTableAdapter.Update(this.mainDBDataSet.Workload);
		}
	}
}
