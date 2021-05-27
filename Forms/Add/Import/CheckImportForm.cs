using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CafedralDB.Forms.Add.Import
{
    public partial class CheckImportForm : Form
    {
        public CheckImportForm()
        {
            InitializeComponent();
        }

        private void buttonNo_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Поправьте данные и запустите импорт заново");
            this.Close();
        }
    }
}
