using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

namespace CafedralDB.SourceCode.Model.Exporter
{
    public static class ExportSemester
    {
		public static void Export(string year, Utilities.SemesterName semester)
		{

			string path = System.Windows.Forms.Application.StartupPath + "\\ExcelTemplates\\SemesterTemplate.xltx";
			string semesterParam = Utilities.getSemesterConditionString(semester);
			float FIT_sum = 0.0f, MSF_sum = 0.0f, Mag_Sum = 0.0f, IDPO_sum = 0.0f;
			float FIT_lecSum = 0.0f, MSF_lecSum = 0.0f, Mag_lecSum = 0.0f, IDPO_lecSum = 0.0f;
			float contractSum = 0.0f, indPlanSum = 0.0f;

			Dictionary<string, int> counts = new Dictionary<string, int>() { { "ФИТ", 0 }, { "МСФ", 0 }, { "ИДПО", 0 }, { "МАГ", 0 } };
			int rowCounter = 0;

			

			string strSQL = "SELECT Speciality.Descr, Discipline.Descr, Group.StudentCount, Semester.WeekCount, Discipline.LectureCount, Workload.ID, Faculty.Descr, Group.EntryYear, Group.Qualification, Group.StudyForm, Discipline.Contr " +
				"FROM StudyYear INNER JOIN((Faculty INNER JOIN Speciality ON Faculty.ID = Speciality.Faculty) " +
				"INNER JOIN(Semester INNER JOIN ([Group] INNER JOIN (Discipline INNER JOIN Workload ON Discipline.ID = Workload.Discipline) ON Group.ID = Workload.Group) ON " +
				"Semester.ID = Workload.Semester) ON Speciality.ID = Group.Speciality) ON(StudyYear.ID = Workload.StudyYear) " +
				"WHERE ((StudyYear.StudyYear) =[@year]) AND " + semesterParam;
			
			/*		sql new
			0 - Speciality.Descr, 
			1 - Discipline.Descr, 
			2 - Group.StudentCount, 
			3 - Semester.WeekCount,
			4 - Discipline.LectureCount, 
			5 - Workload.ID, 
			6 - Faculty.Descr, 
			7 - Group.EntryYear, 
			8 - Group.Qualification, 
			9 - Group.StudyForm, 
			10 - Discipline.Contr - not use
			 */

			DataManager.SharedDataManager();
			var cn = new OleDbConnection(DataManager.Connection.ConnectionString);
			var cmd = new OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = strSQL;
			cmd.Parameters.AddWithValue("@year", year);
			OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			if (reader.HasRows)
			{
				Excel.Application ObjExcel = new Excel.Application();
				Excel.Workbook ObjWorkBook;
				Excel.Worksheet ObjWorkSheet;
				//Книга.
				ObjWorkBook = ObjExcel.Workbooks.Add(path);
				ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
				ObjWorkBook.Title = string.Format("Отчет {0} семестр({1} / {2})", semester, year, Convert.ToInt32(year) + 1);
				ObjWorkSheet.Cells[3, 1] = string.Format("экзаменационной сессии  за  {0} семестр {1} / {2} учебного года", semester, year, Convert.ToInt32(year) + 1);
				while (reader.Read())
				{
					var group = reader[0].ToString();
					var disciplineName = reader[1].ToString();
					var studentCount = Convert.ToInt32(reader[2]);
					var weekCount = Convert.ToInt32(reader[3]);
					var lectureCount = Convert.ToInt32(reader[4]);
					var workloadId = Convert.ToInt32(reader[5]);
					var facultyName = reader[6].ToString();
					var entryYear = DataManager.SharedDataManager().GetStudyYear( Convert.ToInt32(reader[7])).Year;
					var qualification = reader[8].ToString();
					var studyForm = reader[9].ToString();
					var isContract = true;
					var cource = Convert.ToInt32(year) - entryYear + 1;
					var workloadTotal = Calculator.GetWorkloadTotalCost(workloadId);
					var lectureFact = lectureCount * weekCount;

					rowCounter = ApplicationSettings.ExportSettings.SemesterSetting.FITStartRow;
					var workloadAssigns=DataManager.SharedDataManager().GetWorkloadAssigns(workloadId);
                    if (workloadAssigns.Count != 0)
                    {
						foreach(var assign in workloadAssigns)
                        {
							isContract = isContract && assign.IsContract;
                        }
                    }
                    else
                    {
						isContract = false;
                    }
					int index = 0;
					// Если форма обучения не очная
					if (Convert.ToInt32(studyForm) != 1)
					{
						rowCounter = counts["ФИТ"] + counts["ИДПО"] + counts["МСФ"] + ApplicationSettings.ExportSettings.SemesterSetting.IDPOStartRow;
						counts["ИДПО"]++;
						index = counts["ИДПО"];
						IDPO_sum += workloadTotal;
						IDPO_lecSum += lectureFact;

					}
					else if (facultyName == "ФИТ")
					{
						if (Convert.ToInt32(qualification) != 3)
						{
							rowCounter += counts["ФИТ"];
							counts["ФИТ"]++;
							index = counts["ФИТ"];
							FIT_sum += workloadTotal;
							FIT_lecSum += lectureFact;
						}
						else
						{
							rowCounter = counts["ФИТ"] + counts["ИДПО"] + counts["МАГ"] + counts["МСФ"] + ApplicationSettings.ExportSettings.SemesterSetting.MAGStartRow;
							counts["МАГ"]++;
							index = counts["МАГ"];
							Mag_Sum += workloadTotal;
							Mag_lecSum += lectureFact;

						}
					}
					else if (facultyName == "МСФ")
					{
						rowCounter = counts["ФИТ"] + counts["МСФ"] + ApplicationSettings.ExportSettings.SemesterSetting.MSFStartRow;
						counts["МСФ"]++;
						index = counts["МСФ"];
						MSF_sum += workloadTotal;
						MSF_lecSum += lectureFact;

					}

					Excel.Range line = (Excel.Range)ObjWorkSheet.Rows[rowCounter];
					line.Insert();

					ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.IndexColumn] = index;
					ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.GroupColumn] = group;
					ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.CourceColumn] = cource;
					ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.DisciplineColumn] = disciplineName;
					ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.DisciplineCostColumn] = Math.Round(workloadTotal, 2);
					ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.LecFactColumn] = lectureFact;
					ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.StudentsColumn] = studentCount;
					ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.ContractColumn] = isContract ? "Договор" : "Инд.план";
					ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.DisciplineColumn].EntireRow.AutoFit(); //применить автовысоту

					if (isContract)
						contractSum += workloadTotal;
					else
						indPlanSum += workloadTotal;

					rowCounter++;
				}
				rowCounter = counts["ФИТ"] + 10;
				ObjWorkSheet.Cells[rowCounter, 5] = Math.Round(FIT_sum, 2);
				ObjWorkSheet.Cells[rowCounter, 6] = FIT_lecSum;

				rowCounter = counts["ФИТ"] + counts["МСФ"] + counts["ИДПО"] + counts["МАГ"] + 33;
				ObjWorkSheet.Cells[rowCounter, 5] = Math.Round(Mag_Sum, 2);
				ObjWorkSheet.Cells[rowCounter, 6] = Mag_lecSum;

				rowCounter = counts["ФИТ"] + counts["МСФ"] + counts["ИДПО"] + 25;
				ObjWorkSheet.Cells[rowCounter, 5] = Math.Round(IDPO_sum, 2);
				ObjWorkSheet.Cells[rowCounter, 6] = IDPO_lecSum;

				rowCounter = counts["ФИТ"] + counts["МСФ"] + 17;
				ObjWorkSheet.Cells[rowCounter, 5] = Math.Round(MSF_sum, 2);
				ObjWorkSheet.Cells[rowCounter, 6] = MSF_lecSum;


				rowCounter = counts["ФИТ"] + counts["ИДПО"] + counts["МАГ"] + counts["МСФ"] + 38;
				ObjWorkSheet.Cells[rowCounter, 5] = Math.Round(indPlanSum, 2);
				ObjWorkSheet.Cells[rowCounter + 1, 5] = Math.Round(contractSum, 2);
				ObjWorkSheet.Cells[rowCounter + 2, 5] = Math.Round(indPlanSum + contractSum, 2);

				ObjExcel.Visible = true;
				ObjExcel.UserControl = true;
			}
			cn.Close();
		}
	}
}
