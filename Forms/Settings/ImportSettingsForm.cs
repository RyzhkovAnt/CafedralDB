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
using CafedralDB.SourceCode.Settings;

namespace CafedralDB.Forms.Settings
{
	public partial class ImportSettingsForm : Form
	{
		ImportSetting setting;
		Dictionary<Field, object> changedValue;
		public ImportSettingsForm()
		{
			InitializeComponent();
			setting = new ImportSetting();
			changedValue = new Dictionary<Field, object>();
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

		void OnChangeField(Field field,string value)
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
			OnChangeField(ImportSettingFields.StartReadingRow, textBoxStartReadingRow.Text);
        }

        private void textBoxGroupColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField(ImportSettingFields.GroupColumn, textBoxGroupColumn.Text);
		}

        private void textBoxSemesterColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField(ImportSettingFields.SemesterColumn, textBoxSemesterColumn.Text);
		}

        private void textBoxWeekColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField(ImportSettingFields.WeeksColumn, textBoxWeekColumn.Text);
		}

        private void textBoxDisciplineNameColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField(ImportSettingFields.DisciplineNameColumn, textBoxDisciplineNameColumn.Text);
		}

        private void textBoxLecturesColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField(ImportSettingFields.LecturesColumn, textBoxLecturesColumn.Text);
		}

        private void textBoxLabsColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField(ImportSettingFields.LabsColumn, textBoxLabsColumn.Text);
		}

        private void textBoxPracticesColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField(ImportSettingFields.PracticesColumn, textBoxPracticesColumn.Text);
		}

		private void textBoxKzColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField(ImportSettingFields.KzColumn, textBoxKzColumn.Text);
		}

		private void textBoxKrColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField(ImportSettingFields.KrColumn, textBoxKrColumn.Text);
		}

		private void textBoxKpColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField(ImportSettingFields.KpColumn, textBoxKpColumn.Text);
		}

		private void textBoxEkzColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField(ImportSettingFields.EkzColumn, textBoxEkzColumn.Text);
		}

		private void textBoxZachColumn_TextChanged(object sender, EventArgs e)
        {
			OnChangeField(ImportSettingFields.ZachColumn, textBoxZachColumn.Text);
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
