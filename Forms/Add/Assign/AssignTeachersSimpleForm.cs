using Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CafedralDB.Forms.Add.Assign
{
    public partial class AssignTeachersSimpleForm : Form
    {
        public AssignTeachersSimpleForm()
        {
            InitializeComponent();
        }

        private void AssignTeachersSimpleForm_Load(object sender, EventArgs e)
        {
			// TODO: This line of code loads data into the 'mainDBDataSet.StudyYear' table. You can move, or remove it, as needed.
			this.studyYearTableAdapter.Fill(this.mainDBDataSet.StudyYear);

            RefillTable();
			UpdateSumms();
      
        }

        private void buttonAssign_Click(object sender, EventArgs e)
		{
			if (dataGridViewFreeTeachers.SelectedRows.Count == 1)
			{
				int teacherID = Convert.ToInt32(dataGridViewFreeTeachers.SelectedRows[0].Cells[0].Value);
				int workloadID = this.mainDBDataSet.DataTableForAssigns.ElementAt(dataGridView1.SelectedRows[0].Index).ID;
				int studCount = Convert.ToInt32(textBoxStudentCount.Text);
				DataManager.SharedDataManager().AssignTeacher(workloadID, teacherID, studCount);
				FindAssigns(workloadID);
				UpdateSumms();
			}
		}
		private void buttonFree_Click(object sender, EventArgs e)
		{
			if (dataGridViewAssignedTeachers.SelectedRows.Count == 1)
			{
				int teacherID = Convert.ToInt32(dataGridViewAssignedTeachers.SelectedRows[0].Cells[0].Value);
				int workloadID = this.mainDBDataSet.DataTableForAssigns.ElementAt(dataGridView1.SelectedRows[0].Index).ID;
				DataManager.SharedDataManager().FreeTeacher(workloadID, teacherID);
				FindAssigns(workloadID);
				UpdateSumms();
			}
		}

		private void dataGridView1_SelectionChanged(object sender, EventArgs e)
		{
			//найти выбранную из датагрида нагрузку
			//обновить таблицы свободных и назначенных преподов
			//подсчитать всего часов сколько

			if (dataGridView1.SelectedRows.Count == 1)
			{
				int workloadID = this.mainDBDataSet.DataTableForAssigns.ElementAt(dataGridView1.SelectedRows[0].Index).ID;
				Model.Entity.Workload workload = DataManager.SharedDataManager().GetWorkload(workloadID);
				Discipline disc = DataManager.SharedDataManager().GetDiscipline(workload.Discipline);
				Group group = DataManager.SharedDataManager().GetGroup(workload.Group);
				if (disc.TypeID == 2)
					textBoxStudentCount.Text = "0";
				else
					textBoxStudentCount.Text = group.StudentCount.ToString();
				FindAssigns(workloadID);

				textBoxWorkloadHours.Text = Calculator.GetWorkloadTotalCost(workloadID).ToString();
			}
		}

		public void RefillTable()
		{
			try
			{
				//MessageBox.Show(string.Format("{0},{1},{2}", comboBoxStudyYear.SelectedValue,
				//comboBoxStudyYear.SelectedText, comboBoxStudyYear.SelectedItem));
				int yearID;
				if (comboBoxStudyYear.SelectedValue == null)
				{
					comboBoxStudyYear.SelectedItem = comboBoxStudyYear.Items[0];
					yearID = (int)comboBoxStudyYear.SelectedValue;
				}
				else
				{
					yearID = (int)comboBoxStudyYear.SelectedValue;
				}
				
				//this.dataTableForAssignsTableAdapter.ClearBeforeFill = true;

				OleDbCommand command = new System.Data.OleDb.OleDbCommand();
					command.Connection = Model.DataManager.Connection;
					command.CommandText = @"SELECT        Workload.ID, Discipline.Descr AS Дисциплина, [Group].Descr AS Группа, StudyForm.DescrRus AS Форма, Semester.Descr AS Семестр, StudyYear.ID AS ГодID, StudyYear.StudyYear
				FROM (((((Discipline INNER JOIN
                         Workload ON Discipline.ID = Workload.Discipline) INNER JOIN
                         [Group] ON Workload.[Group] = [Group].ID) INNER JOIN
                         Semester ON Workload.Semester = Semester.ID) INNER JOIN
                         StudyForm ON [Group].StudyForm = StudyForm.ID) INNER JOIN
                         StudyYear ON Workload.StudyYear = StudyYear.ID)
				WHERE (StudyYear.ID = [@param])";


					command.CommandType = global::System.Data.CommandType.Text;
					Model.DataManager.SharedDataManager();
					//Model.DataManager.Connection.Open();
					//this.dataTableForAssignsTableAdapter.Connection = Model.DataManager.Connection;
					command.Parameters.AddWithValue("@param", yearID);
					this.dataTableForAssignsTableAdapter.Adapter.SelectCommand = command;
					this.dataTableForAssignsTableAdapter.Adapter.SelectCommand.Connection = Model.DataManager.Connection;
					this.mainDBDataSet.DataTableForAssigns.Clear();
					//dataTableForAssignsTableAdapter.Adapter.SelectCommand.Parameters.Clear();
					//dataTableForAssignsTableAdapter.Adapter.SelectCommand.Parameters.AddWithValue("@param", yearID);

					this.dataTableForAssignsTableAdapter.Adapter.Fill(this.mainDBDataSet.DataTableForAssigns);

					
					//Model.DataManager.Connection.Close();
			}
			catch (System.Exception ex)
			{
				System.Windows.Forms.MessageBox.Show(ex.Message);
			}

			
		}

		public void FindAssigns(int workloadID)
		{
			List<WorkloadAssign> assigns = DataManager.SharedDataManager().GetWorkloadAssigns(workloadID);

			List<Employee> all = DataManager.SharedDataManager().GetEmployees();
			List<Employee> assigned = new List<Employee>();

			for(int i=0;i<assigns.Count;i++)
			{
				int teacherID = assigns[i].EmployeeID;

				Employee foundedAssign = all.Find(x => x.ID == teacherID);
				assigned.Add(foundedAssign);
				all.Remove(foundedAssign);
			}
			dataGridViewAssignedTeachers.Rows.Clear();
			dataGridViewFreeTeachers.Rows.Clear();

			//dataGridViewFreeTeachers.RowCount = all.Count;
			//dataGridViewAssignedTeachers.RowCount = assigned.Count;

			foreach (Employee empl in all)
			{
				//listBoxFreeTeachers.Items.Add(empl.Name);
				object[] o = { empl.ID, empl.Name };
				dataGridViewFreeTeachers.Rows.Add(o);
			}

			foreach (Employee empl in assigned)
			{
				//listBoxAssignedTeachers.Items.Add(empl.Name);
				object[] o = { empl.ID, empl.Name };
				dataGridViewAssignedTeachers.Rows.Add(o);
			}
		}

		public void UpdateSumms()
		{
			dataGridViewEmployeeSumms.Rows.Clear();
            List<Employee> employees = DataManager.SharedDataManager().GetEmployees();
            
            foreach (Employee emp in employees)
			{
				object[] o = { emp.Name, Calculator.GetEmployeeAllWorkloadsCost(emp.ID, (int)comboBoxStudyYear.SelectedValue) };
				dataGridViewEmployeeSumms.Rows.Add(o);

				dataGridViewEmployeeSumms.Sort(dataGridViewEmployeeSumms.Columns[1], ListSortDirection.Descending);
			}
			if (dataGridView1.RowCount>0)
			{
				int selectedRow = dataGridView1.SelectedRows[0].Index;
				int workID = Convert.ToInt32(dataGridView1.Rows[selectedRow].Cells[0].Value);
				dataGridView1.Rows[selectedRow].Cells[7].Value = DataManager.SharedDataManager().CheckWorkloadAssign(workID);
			}
        }

        private void comboBoxStudyYear_SelectionChangeCommitted(object sender, EventArgs e)
		{
			RefillTable();
            UpdateSumms();
        }

		private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
		{
			for (int i = 0; i < dataGridView1.RowCount; i++)
			{
				int workID = Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value);
				dataGridView1.Rows[i].Cells[6].Value = Calculator.GetWorkloadTotalCost(workID);
				dataGridView1.Rows[i].Cells[7].Value = DataManager.SharedDataManager().CheckWorkloadAssign(workID);
			}

			dataGridView1.Sort(dataGridView1.Columns[5], ListSortDirection.Ascending);
		}
	}
}
