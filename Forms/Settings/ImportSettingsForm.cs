using ApplicationSettings;
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
	public partial class ImportSettingsForm : Form
	{
		public ImportSettingsForm()
		{

			ImportSettings.FromRegistry();
			InitializeComponent();
			textBoxStartReadingRow.Text = ImportSettings.StartReadingRow.ToString();
			textBoxGroupColumn.Text = ImportSettings.GroupColumn.ToString();
			textBoxSemesterColumn.Text = ImportSettings.SemesterColumn.ToString();
			textBoxWeekColumn.Text = ImportSettings.WeeksColumn.ToString();
			textBoxDisciplineNameColumn.Text = ImportSettings.DisciplineNameColumn.ToString();
			textBoxLecturesColumn.Text = ImportSettings.LecturesColumn.ToString();
			textBoxLabsColumn.Text = ImportSettings.LabsColumn.ToString();
			textBoxPracticesColumn.Text = ImportSettings.PracticesColumn.ToString();
			textBoxKzColumn.Text = ImportSettings.KzColumn.ToString();
			textBoxKrColumn.Text = ImportSettings.KrColumn.ToString();
			textBoxKpColumn.Text = ImportSettings.KpColumn.ToString();
			textBoxEkzColumn.Text = ImportSettings.EkzColumn.ToString();
			textBoxZachColumn.Text = ImportSettings.ZachColumn.ToString();
			textBoxOtherColumn.Text = ImportSettings.OtherColumn.ToString();
            textBoxContractColumn.Text = ImportSettings.ContractColumn.ToString();
		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			ImportSettings.StartReadingRow = Convert.ToInt32(textBoxStartReadingRow.Text);
			ImportSettings.GroupColumn = Convert.ToInt32(textBoxGroupColumn.Text);
			ImportSettings.SemesterColumn = Convert.ToInt32(textBoxSemesterColumn.Text);
			ImportSettings.WeeksColumn = Convert.ToInt32(textBoxWeekColumn.Text);
			ImportSettings.DisciplineNameColumn = Convert.ToInt32(textBoxDisciplineNameColumn.Text);
			ImportSettings.LecturesColumn = Convert.ToInt32(textBoxLecturesColumn.Text);
			ImportSettings.LabsColumn = Convert.ToInt32(textBoxLabsColumn.Text);
			ImportSettings.PracticesColumn = Convert.ToInt32(textBoxPracticesColumn.Text);
			ImportSettings.KzColumn = Convert.ToInt32(textBoxKzColumn.Text);
			ImportSettings.KrColumn = Convert.ToInt32(textBoxKrColumn.Text);
			ImportSettings.KpColumn = Convert.ToInt32(textBoxKpColumn.Text);
			ImportSettings.EkzColumn = Convert.ToInt32(textBoxEkzColumn.Text);
			ImportSettings.ZachColumn = Convert.ToInt32(textBoxZachColumn.Text);
			ImportSettings.OtherColumn = Convert.ToInt32(textBoxOtherColumn.Text);
            ImportSettings.ContractColumn = Convert.ToInt32(textBoxContractColumn.Text);
			MessageBox.Show("Сохранено");
			ImportSettings.ToRegistry();
		}
	}
}
