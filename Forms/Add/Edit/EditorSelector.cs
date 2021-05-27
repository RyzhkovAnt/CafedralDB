using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CafedralDB.Forms.Add.Edit
{
	public partial class EditorSelector : Form
	{
		public EditorSelector()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			Form editor = new Form();
			switch (this.comboBoxTableSelection.SelectedIndex)
			{
				case 0:
				{
						editor = new EditTables.AcademicDegreeEditor();
						break;
				}
				case 1:
					{
						editor = new EditTables.AcademicRankEditor();
						break;
					}
				case 2:
					{
						editor = new EditTables.DepartmentEditor();
						break;
					}
				case 3:
					{
						editor = new EditTables.DisciplineEditor();
						break;
					}
				case 4:
					{
						editor = new EditTables.DisciplineTypeEditor();
						break;
					}
				case 5:
					{
						editor = new EditTables.EmployeeEditor();
						break;
					}
				case 6:
					{
						editor = new EditTables.FacultyEditor();
						break;
					}
				case 7:
					{
						editor = new EditTables.GroupEditor();
						break;
					}
				case 8:
					{
						editor = new EditTables.QualificationEditor();
						break;
					}
				case 9:
					{
						editor = new EditTables.SemesterEditor();
						break;
					}
				case 10:
					{
						editor = new EditTables.SpecialityEditor();
						break;
					}
				case 11:
					{
						editor = new EditTables.StudentEditor();
						break;
					}
				case 12:
					{
						editor = new EditTables.StudyFormEditor();
						break;
					}
				case 13:
					{
						//editor = new EditTables.StudyYearEditor();
						break;
					}
				case 14:
					{
						//editor = new EditTables.WorkingPositionEditor();
						break;
					}
				case 15:
					{
						editor = new EditTables.WorkloadEditor();
						break;
					}
				case 16:
					{
						editor = new EditTables.WorkloadAssignEditor();
						break;
					}
			}
			if(editor!=null)
				editor.Show();
		}
	}
}
