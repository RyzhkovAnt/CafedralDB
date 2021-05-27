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
	public partial class CalculationSettingsForm : Form
	{
		public CalculationSettingsForm()
		{
			InitializeComponent();
			CalculationSettings.FromRegistry();

			textBoxLectureCost.Text = CalculationSettings.LectureCost.ToString();
			textBoxLabCost.Text= CalculationSettings.LabCost.ToString();
			textBoxPracticeCost.Text= CalculationSettings.PracticeCost.ToString();
			textBoxKonsCost.Text= CalculationSettings.KonsCost.ToString();
			textBoxEkzCost.Text= CalculationSettings.EkzCost.ToString();
			textBoxKRCost.Text= CalculationSettings.KRCost.ToString();
			textBoxKPCost.Text= CalculationSettings.KPCost.ToString();
			textBoxZachCost.Text= CalculationSettings.ZachCost.ToString();
		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			try
			{
				CalculationSettings.LectureCost = Convert.ToSingle(textBoxLectureCost.Text.Replace(',', '.'));
				CalculationSettings.LabCost = Convert.ToSingle(textBoxLabCost.Text.Replace(',', '.'));
				CalculationSettings.PracticeCost = Convert.ToSingle(textBoxPracticeCost.Text.Replace(',', '.'));
				CalculationSettings.KonsCost = Convert.ToSingle(textBoxKonsCost.Text.Replace(',', '.'));
				CalculationSettings.EkzCost = Convert.ToSingle(textBoxEkzCost.Text);
				CalculationSettings.KRCost = Convert.ToSingle(textBoxKRCost.Text.Replace(',', '.'));
				CalculationSettings.KPCost = Convert.ToSingle(textBoxKPCost.Text.Replace(',', '.'));
				CalculationSettings.ZachCost = Convert.ToSingle(textBoxZachCost.Text);
				CalculationSettings.ToRegistry();
				MessageBox.Show("Сохранено");
		}
			catch (FormatException exc)
			{
				MessageBox.Show("Вы ввели число в неверном формате!");
			}
}
	}
}
