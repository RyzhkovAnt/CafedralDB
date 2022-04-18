using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using TemplateEngine.Docx;
using Model;

namespace CafedralDB.SourceCode.Model.Exporter
{
    public static class ExportSemester
    {
		public static void Export(string year, string semester)
		{

			string path = System.Windows.Forms.Application.StartupPath + "\\ExcelTemplates\\SemesterTemplate.xltx";
			string semesterParam;
			float FIT_sum = 0.0f, MSF_sum = 0.0f, Mag_Sum = 0.0f, IDPO_sum = 0.0f;
			float FIT_lecSum = 0.0f, MSF_lecSum = 0.0f, Mag_lecSum = 0.0f, IDPO_lecSum = 0.0f;
			float contractSum = 0.0f, indPlanSum = 0.0f;
			if (semester == "Осенний")
			{
				semesterParam = "(Semester.Descr = '1'  OR  Semester.Descr = '3'  OR  Semester.Descr = '5'  OR  Semester.Descr = '7'  OR  Semester.Descr = '9'  OR  Semester.Descr = '11')";
			}
			else
			{
				semesterParam = "(Semester.Descr = '2'  OR  Semester.Descr = '4'  OR  Semester.Descr = '6'  OR Semester.Descr = '8'  OR  Semester.Descr = '10'  OR  Semester.Descr = '12')";
			}

			Dictionary<string, int> counts = new Dictionary<string, int>() { { "ФИТ", 0 }, { "МСФ", 0 }, { "ИДПО", 0 }, { "МАГ", 0 } };
			int rowCounter = 0;

			string strSQL = "SELECT Speciality.Descr AS SpecDescr, StudyYear.StudyYear, Discipline.Descr AS DiscDescr, [Group].StudentCount, Semester.WeekCount, Discipline.LectureCount, Discipline.PracticeCount, " +
						 "Discipline.LabCount, Discipline.KR, Discipline.KP, Discipline.Ekz, Discipline.Zach, Workload.ID, Semester.ID AS SemID, Faculty.Descr, StudyForm.DescrRus, Discipline.ID AS DiscID, StudyYear.StudyYear, [Group].EntryYear, [Group].Qualification, [Group].StudyForm, Discipline.Contr " +
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
					bool contr = Convert.ToBoolean(reader[21]);

					float summ = Calculator.GetWorkloadTotalCost(Convert.ToInt32(reader[12]));

					rowCounter = ApplicationSettings.ExportSettings.SemesterSetting.FITStartRow;

					int index = 0;

					if (Convert.ToInt32(reader[20]) != 1)
					{
						rowCounter = counts["ФИТ"] + counts["ИДПО"] + counts["МСФ"] + ApplicationSettings.ExportSettings.SemesterSetting.IDPOStartRow;
						counts["ИДПО"]++;
						index = counts["ИДПО"];
						IDPO_sum += summ;
						IDPO_lecSum += lec * weeks;

					}
					else
					if (reader[14].ToString() == "ФИТ")
					{
						if (Convert.ToInt32(reader[19]) != 3)
						{
							rowCounter += counts["ФИТ"];
							counts["ФИТ"]++;
							index = counts["ФИТ"];
							FIT_sum += summ;
							FIT_lecSum += lec * weeks;
							//counts["МСФ"]++;
							//counts["МАГ"]++;
							//counts["ИДПО"]++;
						}
						else
						{
							rowCounter = counts["ФИТ"] + counts["ИДПО"] + counts["МАГ"] + counts["МСФ"] + ApplicationSettings.ExportSettings.SemesterSetting.MAGStartRow;
							counts["МАГ"]++;
							index = counts["МАГ"];
							Mag_Sum += summ;
							Mag_lecSum += lec * weeks;

						}
					}
					else
					if (reader[14].ToString() == "МСФ")
					{
						rowCounter = counts["ФИТ"] + counts["МСФ"] + ApplicationSettings.ExportSettings.SemesterSetting.MSFStartRow;
						counts["МСФ"]++;
						index = counts["МСФ"];
						MSF_sum += summ;
						MSF_lecSum += lec * weeks;

					}

					//if (summ > 0)
					{
						Excel.Range line = (Excel.Range)ObjWorkSheet.Rows[rowCounter];
						line.Insert();

						ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.IndexColumn] = index;
						ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.GroupColumn] = group;
						ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.CourceColumn] = cource;
						ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.DisciplineColumn] = discName;
						ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.DisciplineCostColumn] = Math.Round(summ, 2);
						ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.LecFactColumn] = lec * weeks;
						ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.StudentsColumn] = countStud;
						ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.ContractColumn] = contr ? "Договор" : "Инд.план";
						ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.DisciplineColumn].EntireRow.AutoFit(); //применить автовысоту

						if (contr)
							contractSum += summ;
						else
							indPlanSum += summ;
					}

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
