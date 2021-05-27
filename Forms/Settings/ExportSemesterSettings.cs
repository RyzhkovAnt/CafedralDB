using ApplicationSettings.ExportSettings;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CafedralDB.Forms.Settings
{
	public partial class ExportSemesterSettings : Form
	{
		public ExportSemesterSettings()
		{
			InitializeComponent();
			ApplicationSettings.ExportSettings.SemesterSetting.FromRegistry();
			textBoxFITStartRow.Text = SemesterSetting.FITStartRow.ToString();
			textBoxMSFStartRow.Text = SemesterSetting.MSFStartRow.ToString();
			textBoxIDPOStartRow.Text = SemesterSetting.IDPOStartRow.ToString();
			textBoxMAGStartRow.Text= SemesterSetting.MAGStartRow.ToString();
			textBoxIndexColumn.Text= SemesterSetting.IndexColumn.ToString();
			textBoxGroupColumn.Text= SemesterSetting.GroupColumn.ToString();
			textBoxCourceColumn.Text= SemesterSetting.CourceColumn.ToString();
			textBoxDisciplineColumn.Text= SemesterSetting.DisciplineColumn.ToString();
			textBoxDisciplineCostColumn.Text= SemesterSetting.DisciplineCostColumn.ToString();
			textBoxLecFactColumn.Text= SemesterSetting.LecFactColumn.ToString();
			textBoxStudentsColumn.Text= SemesterSetting.StudentsColumn.ToString();
            textBoxContractColumn.Text = SemesterSetting.ContractColumn.ToString();
		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			SemesterSetting.FITStartRow = Convert.ToInt32(textBoxFITStartRow.Text);
			SemesterSetting.MSFStartRow = Convert.ToInt32(textBoxMSFStartRow.Text);
			SemesterSetting.IDPOStartRow = Convert.ToInt32(textBoxIDPOStartRow.Text);
			SemesterSetting.MAGStartRow = Convert.ToInt32(textBoxMAGStartRow.Text);
			SemesterSetting.IndexColumn = Convert.ToInt32(textBoxIndexColumn.Text);
			SemesterSetting.GroupColumn = Convert.ToInt32(textBoxGroupColumn.Text);
			SemesterSetting.CourceColumn = Convert.ToInt32(textBoxCourceColumn.Text);
			SemesterSetting.DisciplineColumn = Convert.ToInt32(textBoxDisciplineColumn.Text);
			SemesterSetting.DisciplineCostColumn = Convert.ToInt32(textBoxDisciplineCostColumn.Text);
			SemesterSetting.LecFactColumn = Convert.ToInt32(textBoxLecFactColumn.Text);
			SemesterSetting.StudentsColumn = Convert.ToInt32(textBoxStudentsColumn.Text);
            SemesterSetting.ContractColumn = Convert.ToInt32(textBoxContractColumn.Text);
			MessageBox.Show("Сохранено");

			SemesterSetting.ToRegistry();
		}
	}
}
