﻿using System;
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
	public partial class SpecialityEditor : Form
	{
		public SpecialityEditor()
		{
			InitializeComponent();
		}

		private void SpecialityEditor_Load(object sender, EventArgs e)
		{
			// TODO: This line of code loads data into the 'mainDBDataSet.Speciality' table. You can move, or remove it, as needed.
			this.specialityTableAdapter.Fill(this.mainDBDataSet.Speciality);

		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			this.specialityTableAdapter.Update(this.mainDBDataSet.Speciality);
		}
	}
}
