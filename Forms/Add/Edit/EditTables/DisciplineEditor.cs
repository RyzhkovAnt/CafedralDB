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
	public partial class DisciplineEditor : Form
	{
		public DisciplineEditor()
		{
			InitializeComponent();
		}

		private void DisciplineEditor_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.Discipline' table. You can move, or remove it, as needed.
			this.disciplineTableAdapter.Fill(this.mainDBDataSet.Discipline);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			disciplineTableAdapter.Update(this.mainDBDataSet.Discipline);
		}

		private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
		{

		}
	}
}
