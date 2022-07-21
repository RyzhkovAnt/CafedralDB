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
	public partial class EmployeeEditor : Form
	{
		public EmployeeEditor()
		{
			InitializeComponent();
		}

		private void EmployeeEditor_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.Employee' table. You can move, or remove it, as needed.
			this.employeeTableAdapter.Fill(this.mainDBDataSet.Employee);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			this.employeeTableAdapter.Update(this.mainDBDataSet.Employee);
		}

        private void employeeBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }
    }
}
