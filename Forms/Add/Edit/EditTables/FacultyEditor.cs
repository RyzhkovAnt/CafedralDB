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
	public partial class FacultyEditor : Form
	{
		public FacultyEditor()
		{
			InitializeComponent();
		}

		private void FacultyEditor_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.Faculty' table. You can move, or remove it, as needed.
			this.facultyTableAdapter.Fill(this.mainDBDataSet.Faculty);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			this.facultyTableAdapter.Update(this.mainDBDataSet.Faculty);
		}
	}
}
