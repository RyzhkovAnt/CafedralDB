using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Model
{
	namespace Exporter
	{
		public static class IndPlan
		{
			public static void ExportIndPlan(int teacherID, string year)
			{
				string path = System.Windows.Forms.Application.StartupPath + "\\ExcelTemplates\\IndPlanTemplate.xltx";

				Employee teacher = DataManager.SharedDataManager().GetEmployee(teacherID);

				Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
				Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
				Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
				//Книга.
				ObjWorkBook = ObjExcel.Workbooks.Add(path);//System.Reflection.Missing.Value);
														   //Таблица.
				ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
				ObjWorkBook.Title = string.Format("Индивидуальный план {0} ({1} / {2})", teacher.Name, year, Convert.ToInt32(year) + 1);
				ObjWorkSheet.Cells[
				ApplicationSettings.ExportSettings.IndPlan.YearsDescriptionRowCell[0],
				 ApplicationSettings.ExportSettings.IndPlan.YearsDescriptionRowCell[1]] = string.Format("На  {0} / {1} учебный год", year, Convert.ToInt32(year) + 1);
				ObjWorkSheet.Cells[ApplicationSettings.ExportSettings.IndPlan.TeacherFIORowCell[0],
				 ApplicationSettings.ExportSettings.IndPlan.TeacherFIORowCell[1]] = teacher.Name;

				PrintSemester(teacherID, "Осенний", year, ObjWorkBook);
				PrintSemester(teacherID, "Весенний", year, ObjWorkBook);

				ObjExcel.Visible = true;
				ObjExcel.UserControl = true;

			}

			private static void PrintSemester(int teacherID, string semester, string year, Workbook objWorkBook)
			{
				string semesterParam;
				int DisciplineDescrList;
				int DisciplineParamsList;

				if (semester == "Осенний")
				{
					semesterParam = "(Semester.Descr = '1'  OR  Semester.Descr = '3'  OR  Semester.Descr = '5'  OR  Semester.Descr = '7'  OR  Semester.Descr = '9'  OR  Semester.Descr = '11')";
					DisciplineDescrList = 2;
					DisciplineParamsList = 4;
				}
				else
				{
					semesterParam = "(Semester.Descr = '2'  OR  Semester.Descr = '4'  OR  Semester.Descr = '6'  OR Semester.Descr = '8'  OR  Semester.Descr = '10'  OR  Semester.Descr = '12')";
					DisciplineDescrList = 3;
					DisciplineParamsList = 5;
				}

				string strSQL = "SELECT Discipline.Descr, Faculty.Descr AS Expr1, Speciality.Descr AS Expr2, [Group].EntryYear, [Group].StudentCount, Semester.WeekCount, Discipline.LectureCount, Discipline.PracticeCount, Discipline.LabCount, " +
						 "Discipline.KR, Discipline.KP, Discipline.RGR, Discipline.Zach, Discipline.Ekz, Discipline.Kons, DisciplineType.Descr AS Expr3, StudyYear.StudyYear, WorkloadAssign.Teacher " +
						"FROM(((StudyYear INNER JOIN " +
						 "((Semester INNER JOIN " +
						 "((DisciplineType INNER JOIN " +
						 "Discipline ON DisciplineType.ID = Discipline.DisciplineType) INNER JOIN " +
						 "Workload ON Discipline.ID = Workload.Discipline) ON Semester.ID = Workload.Semester) INNER JOIN " +
						 "((Speciality INNER JOIN " +
						 "Faculty ON Speciality.Faculty = Faculty.ID) INNER JOIN " +
						 "[Group] ON Speciality.ID = [Group].Speciality) ON Workload.[Group] = [Group].ID) ON StudyYear.ID = Workload.StudyYear) INNER JOIN " +
						"WorkloadAssign ON Workload.ID = WorkloadAssign.Workload) INNER JOIN " +
						"Qualification ON[Group].Qualification = Qualification.ID) " +
						"WHERE(WorkloadAssign.Teacher = @param1) AND(StudyYear.StudyYear = @param2) AND " + semesterParam;

				DataManager.SharedDataManager();
				var cn = new System.Data.OleDb.OleDbConnection(DataManager.Connection.ConnectionString);
				var cmd = new System.Data.OleDb.OleDbCommand();
				cn.Open();
				cmd.Connection = cn;
				cmd.CommandType = CommandType.Text;
				cmd.CommandText = strSQL;
				cmd.Parameters.AddWithValue("@param1", teacherID);
				cmd.Parameters.AddWithValue("@param2", year);
				System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();


				if (reader.HasRows)
				{
					Worksheet DescrSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.Sheets[DisciplineDescrList];
					Worksheet ParamsSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.Sheets[DisciplineParamsList];
					int currentRow = ApplicationSettings.ExportSettings.IndPlan.DisciplinesStartRow;
					while (reader.Read())
					{
						DescrSheet.Cells[currentRow, ApplicationSettings.ExportSettings.IndPlan.DisciplineNameColumn] = reader[0].ToString();
						int entry = DataManager.SharedDataManager().GetStudyYear(Convert.ToInt32(reader[3])).Year;
						int cource = Convert.ToInt32(year) - entry + 1;
						string descr = string.Format("{0},{1},{2} курс", reader[1], reader[2], cource);
						DescrSheet.Cells[currentRow, ApplicationSettings.ExportSettings.IndPlan.DisciplineDescrColumn] = descr;
						DescrSheet.Cells[currentRow, ApplicationSettings.ExportSettings.IndPlan.StudentsCountColumn] = reader[4];
						int countStud = Convert.ToInt32(reader[4]);
						int weeks = Convert.ToInt32(reader[5]);
						float lecFact = weeks * Convert.ToInt32(reader[6]) * ApplicationSettings.CalculationSettings.LectureCost;
						float pracFact = weeks * Convert.ToInt32(reader[7]) * ApplicationSettings.CalculationSettings.PracticeCost;
						float labFact = weeks * Convert.ToInt32(reader[8]) * ApplicationSettings.CalculationSettings.LabCost;
						float konsFact = Convert.ToBoolean(reader[14]) ? 0.05f * Convert.ToInt32(reader[6]) * weeks : 0;
						float ekzFact = Convert.ToBoolean(reader[13]) ? 0.33f * countStud + 2 : 0;
						float zachFact = Convert.ToBoolean(reader[12]) ? 0.25f * countStud : 0;
						float KPFact = Convert.ToBoolean(reader[10]) ? ApplicationSettings.CalculationSettings.KPCost * countStud : 0;
						float KRFact = Convert.ToBoolean(reader[9]) ? ApplicationSettings.CalculationSettings.KRCost * countStud : 0;
						ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlan.DisciplineSettingsStartColumn] = lecFact;
						ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlan.DisciplineSettingsStartColumn + 1] = pracFact;
						ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlan.DisciplineSettingsStartColumn + 2] = labFact;
						ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlan.DisciplineSettingsStartColumn + 3] = konsFact;
						ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlan.DisciplineSettingsStartColumn + 4] = ekzFact;
						ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlan.DisciplineSettingsStartColumn + 5] = zachFact;
						ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlan.DisciplineSettingsStartColumn + 6] = KPFact;
						ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlan.DisciplineSettingsStartColumn + 7] = KRFact;
						currentRow += 2;
					}
				}
				cn.Close();
			}
		}
		public static class Semester
		{
			public static void ExportSemester(string year, string semester)
			{

				string path = System.Windows.Forms.Application.StartupPath + "\\ExcelTemplates\\SemesterTemplate.xltx";
				string semesterParam;

				if (semester == "Осенний")
				{
					semesterParam = "(Semester.Descr = '1'  OR  Semester.Descr = '3'  OR  Semester.Descr = '5'  OR  Semester.Descr = '7'  OR  Semester.Descr = '9'  OR  Semester.Descr = '11')";
				}
				else
				{
					semesterParam = "(Semester.Descr = '2'  OR  Semester.Descr = '4'  OR  Semester.Descr = '6'  OR Semester.Descr = '8'  OR  Semester.Descr = '10'  OR  Semester.Descr = '12')";
				}

				Dictionary<string, int> counts = new Dictionary<string, int>() { { "ФИТ", 0 }, { "МСФ", 0 }, { "ИДПО", 0 }, { "МАГ", 0 } };

				string strSQL = "SELECT Speciality.Descr AS SpecDescr, StudyYear.StudyYear, Discipline.Descr AS DiscDescr, [Group].StudentCount, Semester.WeekCount, Discipline.LectureCount, Discipline.PracticeCount, " +
							 "Discipline.LabCount, Discipline.KR, Discipline.KP, Discipline.Ekz, Discipline.Zach, Workload.ID, Semester.ID AS SemID, Faculty.Descr, StudyForm.DescrRus, Discipline.ID AS DiscID, StudyYear.StudyYear, [Group].EntryYear, [Group].Qualification, [Group].StudyForm " +
							"FROM(((((((Speciality INNER JOIN " +
							 "[Group] ON Speciality.ID = [Group].Speciality) INNER JOIN " +
							 "Faculty ON Speciality.Faculty = Faculty.ID) INNER JOIN " +
							 "StudyForm ON[Group].StudyForm = StudyForm.ID) INNER JOIN " +
							 "Workload ON[Group].ID = Workload.[Group]) INNER JOIN " +
							 "Semester ON Workload.Semester = Semester.ID) INNER JOIN " +
							 "Discipline ON Workload.Discipline = Discipline.ID) INNER JOIN " +
							 "StudyYear ON Workload.StudyYear = StudyYear.ID)" +
							"WHERE(StudyYear.StudyYear = @year) AND " + semesterParam;
				DataManager.SharedDataManager();
				var cn = new System.Data.OleDb.OleDbConnection(DataManager.Connection.ConnectionString);
				var cmd = new System.Data.OleDb.OleDbCommand();
				cn.Open();
				cmd.Connection = cn;
				cmd.CommandType = CommandType.Text;
				cmd.CommandText = strSQL;
				cmd.Parameters.AddWithValue("@year", year);
				System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
				// while there is another record present
				if (reader.HasRows)
				{
					Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
					Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
					Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
					//Книга.
					ObjWorkBook = ObjExcel.Workbooks.Add(path);//System.Reflection.Missing.Value);
															   //Таблица.
					ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
					ObjWorkBook.Title = string.Format("Отчет {0} семестр({1} / {2})", semester, year, Convert.ToInt32(year) + 1);
					ObjWorkSheet.Cells[3, 1] = string.Format("экзаменационной сессии  за  {0} семестр {1} / {2} учебного года", semester, year, Convert.ToInt32(year) + 1);
					while (reader.Read())
					{
						string group = reader[0].ToString();
						int entry = DataManager.SharedDataManager().GetStudyYear(Convert.ToInt32(reader[18])).Year;
						int cource = Convert.ToInt32(year) - entry + 1;
						string discName = reader[2].ToString();
						int countStud = Convert.ToInt32(reader[3]);
						int weeks = Convert.ToInt32(reader[4]);
						int lec = Convert.ToInt32(reader[5]);
						int prac = Convert.ToInt32(reader[6]);
						int lab = Convert.ToInt32(reader[7]);
						bool kr = Convert.ToBoolean(reader[8]);
						bool kp = Convert.ToBoolean(reader[9]);
						bool ekz = Convert.ToBoolean(reader[10]);

						bool zach = Convert.ToBoolean(reader[11]);

						float summ = Calculator.GetWorkloadCost(Convert.ToInt32(reader[12]));

						int rowCounter = ApplicationSettings.ExportSettings.Semester.FITStartRow;

						int index = 0;

						if (Convert.ToInt32(reader[20]) != 1)
						{
							rowCounter = counts["ФИТ"] + counts["ИДПО"] + counts["МСФ"] + ApplicationSettings.ExportSettings.Semester.IDPOStartRow;
							counts["ИДПО"]++;
							index = counts["ИДПО"];
						}
						else
						if (reader[14].ToString() == "ФИТ")
						{
							if (Convert.ToInt32(reader[19]) != 3)
							{
								rowCounter += counts["ФИТ"];
								counts["ФИТ"]++;
								index = counts["ФИТ"];
								//counts["МСФ"]++;
								//counts["МАГ"]++;
								//counts["ИДПО"]++;
							}
							else
							{
								rowCounter = counts["ФИТ"] + counts["ИДПО"] + counts["МАГ"] + counts["МСФ"] + ApplicationSettings.ExportSettings.Semester.MAGStartRow;
								counts["МАГ"]++;
								index = counts["МАГ"];
							}
						}
						else
						if (reader[14].ToString() == "МСФ")
						{
							rowCounter = counts["ФИТ"] + counts["МСФ"] + ApplicationSettings.ExportSettings.Semester.MSFStartRow;
							counts["МСФ"]++;
							index = counts["МСФ"];
						}

						if (summ > 0)
						{
							Range line = (Range)ObjWorkSheet.Rows[rowCounter];
							line.Insert();
							ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.Semester.IndexColumn] = index;
							ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.Semester.GroupColumn] = group;
							ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.Semester.CourceColumn] = cource;
							ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.Semester.DisciplineColumn] = discName;
							ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.Semester.DisciplineCostColumn] = summ;
							ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.Semester.LecFactColumn] = lec * weeks;
							ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.Semester.StudentsColumn] = countStud;
						}

						rowCounter++;
					}

					ObjExcel.Visible = true;
					ObjExcel.UserControl = true;
				}
				cn.Close();
			}
		}
		public static class Workload
		{
			public static void ExportWorkload(string year)
			{
				string path = System.Windows.Forms.Application.StartupPath + "\\ExcelTemplates\\WorkloadTemplate.xltx";

				Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
				Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
				Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
				//Книга.
				ObjWorkBook = ObjExcel.Workbooks.Add(path);//System.Reflection.Missing.Value);
														   //Таблица.
				ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
				ObjWorkBook.Title = string.Format("Нагрузка ({0} / {1})", year, Convert.ToInt32(year) + 1);
				//

				for (int i = 1; i < Math.Min(ObjWorkBook.Sheets.Count, 18); i++)
				{
					ObjWorkBook.Sheets[i].Cells[1, 17] = string.Format("{0} / {1}", year, Convert.ToInt32(year) + 1);
				}

				string strSQL = "SELECT        StudyYear.StudyYear, Workload.ID, StudyForm.ID AS Expr2, Qualification.ID AS Expr1, Discipline.Descr, Speciality.Descr AS Expr4, Faculty.Descr AS Expr3, Semester.Descr AS Expr5, [Group].StudentCount, " +
						 "Semester.WeekCount, Discipline.LectureCount, Discipline.PracticeCount, Discipline.LabCount, Discipline.Ekz, Discipline.Zach, Discipline.RGR, Discipline.KR, Discipline.KP " +
						 "FROM((((((((StudyYear INNER JOIN " +
						 "Workload ON StudyYear.ID = Workload.StudyYear) INNER JOIN " +
						 "Discipline ON Workload.Discipline = Discipline.ID) INNER JOIN " +
						 "[Group] ON Workload.[Group] = [Group].ID) INNER JOIN " +
						 "Qualification ON[Group].Qualification = Qualification.ID) INNER JOIN " +
						 "StudyForm ON[Group].StudyForm = StudyForm.ID) INNER JOIN " +
						 "Speciality ON[Group].Speciality = Speciality.ID) INNER JOIN " +
						 "Faculty ON Speciality.Faculty = Faculty.ID) INNER JOIN " +
						 "Semester ON Workload.Semester = Semester.ID) " +
						 "WHERE(StudyYear.StudyYear = @year) ";
				DataManager.SharedDataManager();
				var cn = new System.Data.OleDb.OleDbConnection(DataManager.Connection.ConnectionString);
				var cmd = new System.Data.OleDb.OleDbCommand();
				cn.Open();
				cmd.Connection = cn;
				cmd.CommandType = CommandType.Text;
				cmd.CommandText = strSQL;
				cmd.Parameters.AddWithValue("@year", year);
				System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

				Dictionary<Worksheet, int> rowCounters = new Dictionary<Worksheet, int>();
				//пишем данные
				for (int i = 2; i <= 17; i++)
				{
					rowCounters.Add(ObjWorkBook.Sheets[i], 6);
				}

				while (reader.Read())
				{
					//Определим листы для записи
					List<Worksheet> toWriteList = new List<Worksheet>();

					if (Convert.ToInt32(reader[2]) == 2)
					{
						toWriteList.Add(ObjWorkBook.Sheets[5]);
					}
					else
					{
						if (reader[6].ToString() == "МСФ")
						{
							toWriteList.Add(ObjWorkBook.Sheets[6]);
						}
						else
						{
							if (Convert.ToInt32(reader[2]) == 1)
							{
								toWriteList.Add(ObjWorkBook.Sheets[4]);
							}
							else
							{
								toWriteList.Add(ObjWorkBook.Sheets[8]);
							}
						}
					}

					List<WorkloadAssign> assigns = DataManager.SharedDataManager().GetWorkloadAssigns(Convert.ToInt32(reader[1]));
					if (assigns.Count != 0)
					{
						foreach (WorkloadAssign assign in assigns)
						{
							string teacherName = DataManager.SharedDataManager().GetEmployee(assign.EmployeeID).Name;
							Worksheet sheet = ObjWorkBook.Sheets[teacherName];
							if (sheet != null)
							{
								toWriteList.Add(sheet);
							}
						}
					}
					else
					{
						if (Convert.ToInt32(reader[3]) == 1)
						{
							toWriteList.Add(ObjWorkBook.Sheets[2]);
						}
						else
						{
							toWriteList.Add(ObjWorkBook.Sheets[3]);
						}
					}

					foreach (Worksheet sheet in toWriteList)
					{

						sheet.Cells[rowCounters[sheet], 1] = rowCounters[sheet] - 5;
						sheet.Cells[rowCounters[sheet], 2] = reader[4].ToString();
						sheet.Cells[rowCounters[sheet], 3] = reader[5].ToString();
						sheet.Cells[rowCounters[sheet], 5] = reader[3].ToString() == "1" ? "о" : "з";
						sheet.Cells[rowCounters[sheet], 4] = reader[6].ToString();
						sheet.Cells[rowCounters[sheet], 6] = reader[7].ToString();
						sheet.Cells[rowCounters[sheet], 7] = reader[8].ToString();
						sheet.Cells[rowCounters[sheet], 8] = reader[9].ToString();
						sheet.Cells[rowCounters[sheet], 11] = reader[10].ToString();
						sheet.Cells[rowCounters[sheet], 12] = reader[11].ToString();
						sheet.Cells[rowCounters[sheet], 13] = reader[12].ToString();
						sheet.Cells[rowCounters[sheet], 13] = reader[12].ToString();
						sheet.Cells[rowCounters[sheet], 14] = Convert.ToBoolean(reader[13]) ? "1" : "";
						sheet.Cells[rowCounters[sheet], 15] = Convert.ToBoolean(reader[14]) ? "+" : "";
						sheet.Cells[rowCounters[sheet], 16] = Convert.ToBoolean(reader[15]) ? "1" : "";
						sheet.Cells[rowCounters[sheet], 17] = Convert.ToBoolean(reader[16]) ? "1" : "";
						sheet.Cells[rowCounters[sheet], 18] = Convert.ToBoolean(reader[17]) ? "1" : "";
						rowCounters[sheet]++;
					}
				}

				//ObjWorkBook.Sheets[1].Cells[1, 1] = ObjWorkBook.Sheets[2].Cells[1, 1] = " ";

				ObjExcel.Visible = true;
				ObjExcel.UserControl = true;
			}
		}
	}
}

