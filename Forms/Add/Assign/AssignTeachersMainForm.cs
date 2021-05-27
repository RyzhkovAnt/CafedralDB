using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CafedralDB.Forms.Add.Assign
{
    public partial class AssignTeachersMainForm : Form
    {
        public AssignTeachersMainForm()
        {
            InitializeComponent();
        }

        private void buttonSimpleAssign_Click(object sender, EventArgs e)
        {
            Assign.AssignTeachersSimpleForm assignForm = new Assign.AssignTeachersSimpleForm();
            assignForm.Show();
        }

        private void buttonSpecialAssign_Click(object sender, EventArgs e)
        {

        }
    }
}
