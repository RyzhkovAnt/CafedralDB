using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CafedralDB.Forms.Add
{
    public partial class UseMainForm : Form
    {
        public UseMainForm()
        {
            InitializeComponent();
        }

        private void buttonAssign_Click(object sender, EventArgs e)
        {

            Assign.AssignTeachersMainForm assignForm = new Assign.AssignTeachersMainForm();
            assignForm.Show();
        }

        private void buttonImport_Click(object sender, EventArgs e)
        {
            Import.ImportForm importForm = new Import.ImportForm();
            importForm.Show();
        }

		private void buttonTables_Click(object sender, EventArgs e)
		{
			Edit.EditorSelector editorsForm = new Edit.EditorSelector();
			editorsForm.Show();
		}
	}
}
