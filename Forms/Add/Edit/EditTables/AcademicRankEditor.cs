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
	public partial class AcademicRankEditor : Form
	{
		public AcademicRankEditor()
		{
			InitializeComponent();
		}

		private void AcademicRankEditor_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.AcademicRank' table. You can move, or remove it, as needed.
			this.academicRankTableAdapter.Fill(this.mainDBDataSet.AcademicRank);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			academicRankTableAdapter.Update(mainDBDataSet.AcademicRank);
		}
	}
}
