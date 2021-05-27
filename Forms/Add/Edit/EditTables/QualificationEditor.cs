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
	public partial class QualificationEditor : Form
	{
		public QualificationEditor()
		{
			InitializeComponent();
		}

		private void QualificationEditor_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.Qualification' table. You can move, or remove it, as needed.
			this.qualificationTableAdapter.Fill(this.mainDBDataSet.Qualification);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			this.qualificationTableAdapter.Update(this.mainDBDataSet.Qualification);
		}
	}
}
