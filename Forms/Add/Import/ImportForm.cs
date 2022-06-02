using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CafedralDB.Forms.Add.Import
{
    public partial class ImportForm : Form
    {
        public ImportForm()
        {
            InitializeComponent();
        }

        private void ImportForm_Load(object sender, EventArgs e)
        {
			// TODO: This line of code loads data into the 'mainDBDataSet1.YearsExistQuery' table. You can move, or remove it, as needed.
			this.yearsExistTableAdapter.Fill(this.mainDBDataSet1.YearsExistQuery);
			// TODO: This line of code loads data into the 'mainDBDataSet1.YearsQuery' table. You can move, or remove it, as needed.
			this.yearsQueryTableAdapter.Fill(this.mainDBDataSet1.YearsQuery);
			// TODO: This line of code loads data into the 'mainDBDataSet.YearsExist' table. You can move, or remove it, as needed.
			this.yearsExistTableAdapter.Fill(this.mainDBDataSet.YearsExistQuery);
            // TODO: This line of code loads data into the 'mainDBDataSet.StudyYear' table. You can move, or remove it, as needed.
            this.studyYearTableAdapter.Fill(this.mainDBDataSet.StudyYear);
            // TODO: This line of code loads data into the 'mainDBDataSet.StudyYear' table. You can move, or remove it, as needed.
            this.studyYearTableAdapter.Fill(this.mainDBDataSet.StudyYear);
        }

        private void buttonImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.FileName = String.Format("Данные.xlsx");
            openDialog.Filter = "Excel files (*.xls*)|*.xls*|All files (*.*)|*.*";
            openDialog.FilterIndex = 1;
           DialogResult res =  openDialog.ShowDialog();

            string filePath = openDialog.FileName;
            string year = comboBoxYear.SelectedValue.ToString();
			if (res == DialogResult.OK || res == DialogResult.Yes)
			{
                SourceCode.Model.Importer.ImportDataFromExcel(filePath, year);
                //SourceCode.Model.Importer.ImportWithGemBox(filePath, year);
                CheckImportForm checkImportForm = new CheckImportForm();
                //checkImportForm.Show();
                MessageBox.Show("Операция завершена!");
			}
        }

        private void fillByExistsYearsToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.studyYearTableAdapter.FillByExistsYears(this.mainDBDataSet.StudyYear);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void fillByExistsYearsToolStripButton1_Click(object sender, EventArgs e)
        {
            try
            {
                this.studyYearTableAdapter.FillByExistsYears(this.mainDBDataSet.StudyYear);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void fillByExistsYearsToolStripButton_Click_1(object sender, EventArgs e)
        {
            try
            {
                this.studyYearTableAdapter.FillByExistsYears(this.mainDBDataSet.StudyYear);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }
    }
}
