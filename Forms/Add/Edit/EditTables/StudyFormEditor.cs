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
	public partial class StudyFormEditor : Form
	{
		public StudyFormEditor()
		{
			InitializeComponent();
		}

		private void StudyFormEditor_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.StudyForm' table. You can move, or remove it, as needed.
			this.studyFormTableAdapter.Fill(this.mainDBDataSet.StudyForm);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			this.studyFormTableAdapter.Update(this.mainDBDataSet.StudyForm);
		}
	}
}
