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
	public partial class GroupEditor : Form
	{
		public GroupEditor()
		{
			InitializeComponent();
		}

		private void GroupEditor_Load(object sender, EventArgs e)
{
			// TODO: This line of code loads data into the 'mainDBDataSet.StudyYear' table. You can move, or remove it, as needed.
			this.studyYearTableAdapter.Fill(this.mainDBDataSet.StudyYear);
			// TODO: This line of code loads data into the 'mainDBDataSet.StudyForm' table. You can move, or remove it, as needed.
			this.studyFormTableAdapter.Fill(this.mainDBDataSet.StudyForm);
			// TODO: This line of code loads data into the 'mainDBDataSet.Qualification' table. You can move, or remove it, as needed.
			this.qualificationTableAdapter.Fill(this.mainDBDataSet.Qualification);
			// TODO: This line of code loads data into the 'mainDBDataSet.Speciality' table. You can move, or remove it, as needed.
			this.specialityTableAdapter.Fill(this.mainDBDataSet.Speciality);
			// TODO: This line of code loads data into the 'mainDBDataSet.Group' table. You can move, or remove it, as needed.
			this.groupTableAdapter.Fill(this.mainDBDataSet.Group);
			// TODO: This line of code loads data into the 'mainDBDataSet.Speciality' table. You can move, or remove it, as needed.
			this.specialityTableAdapter.Fill(this.mainDBDataSet.Speciality);
			// TODO: This line of code loads data into the 'mainDBDataSet.Group' table. You can move, or remove it, as needed.
			this.groupTableAdapter.Fill(this.mainDBDataSet.Group);
			// TODO: This line of code loads data into the 'mainDBDataSet.Speciality' table. You can move, or remove it, as needed.
			this.specialityTableAdapter.Fill(this.mainDBDataSet.Speciality);
			// TODO: This line of code loads data into the 'mainDBDataSet.Group' table. You can move, or remove it, as needed.
			this.groupTableAdapter.Fill(this.mainDBDataSet.Group);

        }

		private void buttonSave_Click(object sender, EventArgs e)
		{
			this.groupTableAdapter.Update(this.mainDBDataSet.Group);
		}
	}
}
