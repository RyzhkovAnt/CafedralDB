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
	public partial class Normas : Form
	{
		public Normas()
		{
			InitializeComponent();
		}

		private void Normas_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.Normas' table. You can move, or remove it, as needed.
			this.normasTableAdapter.Fill(this.mainDBDataSet.Normas);

		}

		private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
		{
			
		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			try
			{
				this.normasTableAdapter.Update(this.mainDBDataSet.Normas);

                ApplicationSettings.CalculationSettings.FromDataBase();
                ApplicationSettings.CalculationSettings.ToRegistry();

            }
			catch (Exception err)
			{
				MessageBox.Show(err.Message);
			}
		}
	}
}
