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
	public partial class DepartmentEditor : Form
	{
		public DepartmentEditor()
		{
			InitializeComponent();
		}

		private void DepartmentEditor_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.Department' table. You can move, or remove it, as needed.
			this.departmentTableAdapter.Fill(this.mainDBDataSet.Department);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			departmentTableAdapter.Update(this.mainDBDataSet.Department);
		}
	}
}
