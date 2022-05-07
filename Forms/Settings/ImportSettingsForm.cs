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
		NewImportSetting setting;
		Dictionary<string, object> changedValue;
		public ImportSettingsForm()
		{
			InitializeComponent();
			setting = new NewImportSetting();
			changedValue = new Dictionary<string, object>();

			textBoxStartReadingRow.Text = setting.StartReadingRow.ToString();
			textBoxGroupColumn.Text = setting.GroupColumn.ToString();
			textBoxSemesterColumn.Text = setting.SemesterColumn.ToString();
			textBoxWeekColumn.Text = setting.WeeksColumn.ToString();
			textBoxDisciplineNameColumn.Text = setting.DisciplineNameColumn.ToString();
			textBoxLecturesColumn.Text = setting.LecturesColumn.ToString();
			textBoxLabsColumn.Text = setting.LabsColumn.ToString();
			textBoxPracticesColumn.Text = setting.PracticesColumn.ToString();
			textBoxKzColumn.Text = setting.KzColumn.ToString();
			textBoxKrColumn.Text = setting.KrColumn.ToString();
			textBoxKpColumn.Text = setting.KpColumn.ToString();
			textBoxEkzColumn.Text = setting.EkzColumn.ToString();
			textBoxZachColumn.Text = setting.ZachColumn.ToString();
		}

		void OnChangeField(string field,string value)
        {
            if (changedValue.ContainsKey(field))
            {
				changedValue[field] = value;
            }
            else
            {
				changedValue.Add(field, value);
            }
        }

		private void buttonSave_Click(object sender, EventArgs e)
		{
			
			
			setting.SaveToRegistry(changedValue);
			MessageBox.Show("Сохранено");
		}

        private void textBoxStartReadingRow_TextChanged(object sender, EventArgs e)
        {
			OnChangeField("StartReadingRow", textBoxStartReadingRow.Text);
        }

        private void textBoxGroupColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField("GroupColumn", textBoxGroupColumn.Text);
		}

        private void textBoxSemesterColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField("SemesterColumn", textBoxSemesterColumn.Text);
		}

        private void textBoxWeekColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField("WeekColumn", textBoxWeekColumn.Text);
		}

        private void textBoxDisciplineNameColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField("DisciplineNameColumn", textBoxDisciplineNameColumn.Text);
		}

        private void textBoxLecturesColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField("LecturesColumn", textBoxLecturesColumn.Text);
		}

        private void textBoxLabsColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField("LabsColumn", textBoxLabsColumn.Text);
		}

        private void textBoxPracticesColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField("PracticesColumn", textBoxPracticesColumn.Text);
		}

		private void textBoxKzColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField("KzColumn", textBoxKzColumn.Text);
		}

		private void textBoxKrColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField("KrColumn", textBoxKrColumn.Text);
		}

		private void textBoxKpColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField("KpColumn", textBoxKpColumn.Text);
		}

		private void textBoxEkzColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField("EkzColumn", textBoxEkzColumn.Text);
		}

		private void textBoxZachColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField("ZachColumn", textBoxZachColumn.Text);
		}

        private void toDefaultButton_Click(object sender, EventArgs e)
        {
			setting.ToDefault();
			textBoxStartReadingRow.Text = setting.StartReadingRow.ToString();
			textBoxGroupColumn.Text = setting.GroupColumn.ToString();
			textBoxSemesterColumn.Text = setting.SemesterColumn.ToString();
			textBoxWeekColumn.Text = setting.WeeksColumn.ToString();
			textBoxDisciplineNameColumn.Text = setting.DisciplineNameColumn.ToString();
			textBoxLecturesColumn.Text = setting.LecturesColumn.ToString();
			textBoxLabsColumn.Text = setting.LabsColumn.ToString();
			textBoxPracticesColumn.Text = setting.PracticesColumn.ToString();
			textBoxKzColumn.Text = setting.KzColumn.ToString();
			textBoxKrColumn.Text = setting.KrColumn.ToString();
			textBoxKpColumn.Text = setting.KpColumn.ToString();
			textBoxEkzColumn.Text = setting.EkzColumn.ToString();
			textBoxZachColumn.Text = setting.ZachColumn.ToString();
		}
    }
}
