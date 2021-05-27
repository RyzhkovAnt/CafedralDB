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
	public partial class WorkloadAssignEditor : Form
	{
		public WorkloadAssignEditor()
		{
			InitializeComponent();
		}

		private void WorkloadAssignEditor_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.Workload' table. You can move, or remove it, as needed.
			this.workloadTableAdapter.Fill(this.mainDBDataSet.Workload);
			// TODO: This line of code loads data into the 'mainDBDataSet.Employee' table. You can move, or remove it, as needed.
			this.employeeTableAdapter.Fill(this.mainDBDataSet.Employee);
			// TODO: This line of code loads data into the 'mainDBDataSet.WorkloadAssign' table. You can move, or remove it, as needed.
			this.workloadAssignTableAdapter.Fill(this.mainDBDataSet.WorkloadAssign);
			// TODO: This line of code loads data into the 'mainDBDataSet.Workload' table. You can move, or remove it, as needed.
			this.workloadTableAdapter.Fill(this.mainDBDataSet.Workload);
			// TODO: This line of code loads data into the 'mainDBDataSet.Employee' table. You can move, or remove it, as needed.
			this.employeeTableAdapter.Fill(this.mainDBDataSet.Employee);
			// TODO: This line of code loads data into the 'mainDBDataSet.WorkoadAssign' table. You can move, or remove it, as needed.
			this.workloadAssignTableAdapter.Fill(this.mainDBDataSet.WorkloadAssign);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			this.workloadAssignTableAdapter.Update(this.mainDBDataSet.WorkloadAssign);
		}
	}
}
