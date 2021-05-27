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
	public partial class WorkingPositionEditor : Form
	{
		public WorkingPositionEditor()
		{
			InitializeComponent();
		}

		private void WorkingPosition_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.WorkingPosition' table. You can move, or remove it, as needed.
			this.workingPositionTableAdapter.Fill(this.mainDBDataSet.WorkingPosition);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			this.workingPositionTableAdapter.Update(this.mainDBDataSet.WorkingPosition);
		}
	}
}
