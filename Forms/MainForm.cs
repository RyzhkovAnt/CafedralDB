using CafedralDB.Forms.Add;
using CafedralDB.Forms.Manage;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CafedralDB
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            Forms.Export.ExportMainForm exportForm = new Forms.Export.ExportMainForm();
            exportForm.Show();
        }

        private void buttonWrite_Click(object sender, EventArgs e)
        {
            UseMainForm addForm = new UseMainForm();
            addForm.Show();
        }

        private void buttonManage_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Эти действия повлияют на расчеты.\r\nВы уверены?", "Подтверждение", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                ManageMainForm manageForm = new ManageMainForm();
                manageForm.Show();
            }
        }

		private void простыеДисциплиныToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Assign.AssignTeachersSimpleForm assignForm = new Forms.Add.Assign.AssignTeachersSimpleForm();
			assignForm.Show();
		}

		private void назначенияНагрузокToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Form editor = new Forms.Add.Edit.EditTables.WorkloadAssignEditor();
			editor.Show();
		}

		private void особыеДисциплиныToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Assign.AssignTeachersSpecialForm assignForm = new Forms.Add.Assign.AssignTeachersSpecialForm();
			assignForm.Show();
		}

		private void семестрToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Export.ExportSemester exportForm = new Forms.Export.ExportSemester();
			exportForm.Show();
		}

		private void импортToolStripMenuItem1_Click(object sender, EventArgs e)
		{
			Forms.Settings.ImportSettingsForm exportForm = new Forms.Settings.ImportSettingsForm();
			exportForm.Show();
		}

		private void настройкиЭкспортаToolStripMenuItem_Click(object sender, EventArgs e)
		{

		}

		private void настройкиРасчетаToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Forms.Settings.CalculationSettingsForm form = new Forms.Settings.CalculationSettingsForm();
			Forms.Add.Edit.EditTables.Normas form = new Forms.Add.Edit.EditTables.Normas();
			form.Show();
		}

		private void семестрToolStripMenuItem1_Click(object sender, EventArgs e)
		{
			Forms.Settings.ExportSemesterSettings exportForm = new Forms.Settings.ExportSemesterSettings();
			exportForm.Show();
		}

		private void ученаяСтепеньToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.AcademicDegreeEditor form = new Forms.Add.Edit.EditTables.AcademicDegreeEditor();
			form.Show();
		}

		private void ученоеЗваниеToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.AcademicRankEditor form = new Forms.Add.Edit.EditTables.AcademicRankEditor();
			form.Show();
		}

		private void кафедрыToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.DepartmentEditor form = new Forms.Add.Edit.EditTables.DepartmentEditor();
			form.Show();
		}

		private void дисциплиныToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.DisciplineEditor form = new Forms.Add.Edit.EditTables.DisciplineEditor();
			form.Show();
		}

		private void типыДисциплинToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.DisciplineTypeEditor form = new Forms.Add.Edit.EditTables.DisciplineTypeEditor();
			form.Show();
		}

		private void сотрудникиToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.EmployeeEditor form = new Forms.Add.Edit.EditTables.EmployeeEditor();
			form.Show();
		}

		private void факультетыToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.FacultyEditor form = new Forms.Add.Edit.EditTables.FacultyEditor();
			form.Show();
		}

		private void группыToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.GroupEditor form = new Forms.Add.Edit.EditTables.GroupEditor();
			form.Show();
		}

		private void квалификацииToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.QualificationEditor form = new Forms.Add.Edit.EditTables.QualificationEditor();
			form.Show();
		}

		private void семестрыToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.SemesterEditor form = new Forms.Add.Edit.EditTables.SemesterEditor();
			form.Show();
		}

		private void специальностиToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.SpecialityEditor form = new Forms.Add.Edit.EditTables.SpecialityEditor();
			form.Show();
		}

		private void студентыToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.StudentEditor form = new Forms.Add.Edit.EditTables.StudentEditor();
			form.Show();
		}

		private void формыОбученияToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.StudyFormEditor form = new Forms.Add.Edit.EditTables.StudyFormEditor();
			form.Show();
		}

		private void учебныеГодыToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.StudyYearEditor form = new Forms.Add.Edit.EditTables.StudyYearEditor();
			form.Show();
		}

		private void должностиToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.WorkingPositionEditor form = new Forms.Add.Edit.EditTables.WorkingPositionEditor();
			form.Show();
		}

		private void нагрузкиToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Edit.EditTables.WorkloadEditor form = new Forms.Add.Edit.EditTables.WorkloadEditor();
			form.Show();
		}

		private void импортToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Import.ImportForm form = new Forms.Add.Import.ImportForm();
			form.Show();
		}

		private void индПланToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Export.ExportIndPlan form = new Forms.Export.ExportIndPlan();
			form.Show();
		}

		private void нагрузкаToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Export.ExportWorkload form = new Forms.Export.ExportWorkload();
			form.Show();
		}

		private void назначениеПреподавателейToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Forms.Add.Assign.AssignTeachersSimpleForm assignForm = new Forms.Add.Assign.AssignTeachersSimpleForm();
			assignForm.Show();
		}

		private void MainForm_Load(object sender, EventArgs e)
		{
            ApplicationSettings.CalculationSettings.FromRegistry();
        }

        private void договорToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Forms.Export.ExportContract form = new Forms.Export.ExportContract();
            form.Show();
        }
    }
}
