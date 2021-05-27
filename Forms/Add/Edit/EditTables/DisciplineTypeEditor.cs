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
	public partial class DisciplineTypeEditor : Form
	{
		public DisciplineTypeEditor()
		{
			InitializeComponent();
		}

		private void DisciplineTypeEditor_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.DisciplineType' table. You can move, or remove it, as needed.
			this.disciplineTypeTableAdapter.Fill(this.mainDBDataSet.DisciplineType);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			disciplineTypeTableAdapter.Update(this.mainDBDataSet.DisciplineType);
		}
	}
}
