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
	public partial class StudentEditor : Form
	{
		public StudentEditor()
		{
			InitializeComponent();
		}

		private void StudentEditor_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.Student' table. You can move, or remove it, as needed.
			this.studentTableAdapter.Fill(this.mainDBDataSet.Student);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			this.studentTableAdapter.Update(this.mainDBDataSet.Student);
		}
	}
}
