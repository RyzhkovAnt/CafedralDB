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
	public partial class SemesterEditor : Form
	{
		public SemesterEditor()
		{
			InitializeComponent();
		}

		private void SemesterEditor_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.Semester' table. You can move, or remove it, as needed.
			this.semesterTableAdapter.Fill(this.mainDBDataSet.Semester);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			semesterTableAdapter.Update(this.mainDBDataSet.Semester);
		}
	}
}
