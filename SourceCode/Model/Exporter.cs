using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using TemplateEngine.Docx;

namespace CafedralDB.SourceCode.Model
{
	namespace Exporter
	{
        [Obsolete("Страрый класс экспорта индивидуального плана")]
		public static class IndPlan
		{
			public static void ExportIndPlan(int teacherID, string year)
			{
				string path = System.Windows.Forms.Application.StartupPath + "\\ExcelTemplates\\IndPlanTemplate.xltx";

				Employee teacher = DataManager.SharedDataManager().GetEmployee(teacherID);

				Excel.Application ObjExcel = new Excel.Application();
				Excel.Workbook ObjWorkBook;
				Excel.Worksheet ObjWorkSheet;
				//Книга.
				ObjWorkBook = ObjExcel.Workbooks.Add(path);//System.Reflection.Missing.Value);
														   //Таблица.
				ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
				ObjWorkBook.Title = string.Format("Индивидуальный план {0} ({1} / {2})", teacher.Name, year, Convert.ToInt32(year) + 1);
				ObjWorkSheet.Cells[
				ApplicationSettings.ExportSettings.IndPlanSetting.YearsDescriptionRowCell[0],
				 ApplicationSettings.ExportSettings.IndPlanSetting.YearsDescriptionRowCell[1]] = string.Format("На  {0} / {1} учебный год", year, Convert.ToInt32(year) + 1);
				ObjWorkSheet.Cells[ApplicationSettings.ExportSettings.IndPlanSetting.TeacherFIORowCell[0],
				 ApplicationSettings.ExportSettings.IndPlanSetting.TeacherFIORowCell[1]] = teacher.Name;

				PrintSemester(teacherID, "Осенний", year, ObjWorkBook);
				PrintSemester(teacherID, "Весенний", year, ObjWorkBook);

				ObjExcel.Visible = true;
				ObjExcel.UserControl = true;

			}

            private static void PrintSemester(int teacherID, string semester, string year, Excel.Workbook objWorkBook)
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

                string strSQL = "SELECT Discipline.Descr, Faculty.Descr AS Expr1, Speciality.Descr AS Expr2, [Group].EntryYear," +
                    " [Group].StudentCount, Semester.WeekCount, DisciplineType.Descr AS Expr3, StudyYear.StudyYear, WorkloadAssign.Teacher," +
                    " Discipline.KonsZaoch, Discipline.GAKPred, Discipline.DPruk, Discipline.DopuskVKR, Discipline.RetzVKR, Discipline.DPretz," +
                    " Discipline.ASPRuk, Discipline.MAGRuk, Discipline.MAGRetz, Discipline.RukKaf, Discipline.NIIR, Discipline.isSpecial," +
                    " Workload.ID FROM (StudyYear INNER JOIN ((Faculty INNER JOIN Speciality ON Faculty.ID = Speciality.Faculty)" +
                    " INNER JOIN (Semester INNER JOIN (Qualification INNER JOIN ((DisciplineType INNER JOIN Discipline ON DisciplineType.ID = Discipline.DisciplineType)" +
                    " INNER JOIN (Workload INNER JOIN [Group] ON Workload.[Group] = Group.ID) ON Discipline.ID = Workload.Discipline) " +
                    "ON Qualification.ID = Group.Qualification) ON Semester.ID = Workload.Semester) ON Speciality.ID = Group.Speciality) " +
                    "ON StudyYear.ID = Workload.StudyYear) INNER JOIN WorkloadAssign ON Workload.ID = WorkloadAssign.Workload " +
                    "WHERE (Discipline.Contr = FALSE) AND (StudyYear.StudyYear=[@param1]) AND (WorkloadAssign.Teacher=[@param2]) AND " + semesterParam;

                DataManager.SharedDataManager();
                var cn = new System.Data.OleDb.OleDbConnection(DataManager.Connection.ConnectionString);
                var cmd = new System.Data.OleDb.OleDbCommand();
                cn.Open();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSQL;
                cmd.Parameters.AddWithValue("@param1", year);
                cmd.Parameters.AddWithValue("@param2", teacherID);
                System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();


                if (reader.HasRows)
                {
                    Excel.Worksheet DescrSheet = (Excel.Worksheet)objWorkBook.Sheets[DisciplineDescrList];
                    Excel.Worksheet ParamsSheet = (Excel.Worksheet)objWorkBook.Sheets[DisciplineParamsList];
                    int currentRow = ApplicationSettings.ExportSettings.IndPlanSetting.DisciplinesStartRow;
                    while (reader.Read())
                    {
                        DescrSheet.Cells[currentRow, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineNameColumn] = reader[0].ToString();
                        int entry = DataManager.SharedDataManager().GetStudyYear(Convert.ToInt32(reader[3])).Year;
                        int cource = Convert.ToInt32(year) - entry + 1;
                        string descr = string.Format("{0},{1},{2} курс", reader[1], reader[2], cource);
                        DescrSheet.Cells[currentRow, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineDescrColumn] = descr;
                        DescrSheet.Cells[currentRow, ApplicationSettings.ExportSettings.IndPlanSetting.StudentsCountColumn] = reader[4];
                        int countStud = Convert.ToInt32(reader[4]);
                        Entity.Workload workload = DataManager.SharedDataManager().GetWorkload(Convert.ToInt32(reader[21]));
                        WorkloadCost workloadCost = Calculator.GetWorkloadCost(workload.ID);
                        CafedralDB.SourceCode.Settings.CalculationSetting calculationSetting = new CafedralDB.SourceCode.Settings.CalculationSetting();

                        if (Convert.ToBoolean(reader[20]))
                        {
                            WorkloadAssign assign = DataManager.SharedDataManager().GetWorkloadAssign(Convert.ToInt32(reader[21]), teacherID);
                            countStud = assign.StudentsCount;
                        }

                        int weeks = Convert.ToInt32(reader[5]);
                        float lecFact = workloadCost.LectureCost;     
                        float pracFact = workloadCost.PracticeCost;  
                        float labFact = workloadCost.LabCost;       
                        float konsFact = workloadCost.KonsCost;
                        float ekzFact = workloadCost.EkzCost;
                        float zachFact = workloadCost.ZachCost;
                        float KPFact = workloadCost.KPCost;     
                        float KRFact = workloadCost.KRCost;    

                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn] = lecFact;
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 1] = pracFact;
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 2] = labFact;
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 3] = konsFact;
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 4] = ekzFact;
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 5] = zachFact;
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 6] = KPFact;
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 7] = KRFact;

                        double other = 0;
                        double geks = 0;
                        double ruk = 0;
                        double uchPr = workloadCost.UchPracCost; 
                        double prPr = workloadCost.PrPracCost;
                        double preddipPr = workloadCost.PredDipPracCost;
                        double NIIR = workloadCost.NIIRCost;
                        double GEK = workloadCost.GEKCost;
                        double GAKpred = workloadCost.GAKPredCost;
                        double GAK = workloadCost.GAKCost;
                        double rukMag = Convert.ToBoolean(reader[16]) ? (30 * Convert.ToInt32(reader[16])) : 0;//рук маг
                        geks = GEK + GAK + GAKpred;

                        ruk += Convert.ToBoolean(reader[11]) ? (countStud * calculationSetting.DPruk) : 0f;//DPruk
                        ruk += Convert.ToBoolean(reader[12]) ? (1 * Convert.ToInt32(reader[4])) : 0f;//DopuskVkr
                        ruk += Convert.ToBoolean(reader[13]) ? (4 * Convert.ToInt32(reader[4])) : 0f;//retzVKR
                        ruk += Convert.ToBoolean(reader[14]) ? (30 * Convert.ToInt32(reader[4])) : 0f;//DPretz
                        ruk += Convert.ToBoolean(reader[15]) ? (calculationSetting.AspRuk * countStud) : 0f;//ASPRuk
                        ruk += rukMag;
                        ruk += Convert.ToBoolean(reader[17]) ? (1 * Convert.ToInt32(reader[4])) : 0f;//MAGRetz
                        ruk += Convert.ToBoolean(reader[18]) ? (1 * Convert.ToInt32(reader[4])) : 0f;//rukKaf

                        other += Convert.ToInt32(reader[19]) != 0 ? (Convert.ToInt32(reader[9]) * 0.5) : 0f;//reader[21]-KonsZaoch
                        other += NIIR;

                        string disciplineName = reader[0].ToString();
                        if (disciplineName.ToLower().Contains("норм") && disciplineName.ToLower().Contains("маг"))
                            other += countStud * calculationSetting.NormocontrolMag;
                        if (disciplineName.ToLower().Contains("доп") && disciplineName.ToLower().Contains("маг"))
                            other += countStud * calculationSetting.DopuskDissMag;

                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 8] = 0;
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 9] = uchPr;//уч пр.
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 10] = prPr;//пр пр.
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 11] = preddipPr;//преддип пр.
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 12] = 0;
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 13] = geks.ToString();
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 14] = ruk.ToString();
                        ParamsSheet.Cells[currentRow - 1, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineSettingsStartColumn + 15] = other.ToString();

                        currentRow += 2;
                    }
                }
                cn.Close();
            }
		}
        [Obsolete("Страрый класс экспорта отчета за семестр")]
		public static class Semester
		{
			public static void ExportSemester(string year, Utilities.SemesterName semester)
			{
                
				string path = System.Windows.Forms.Application.StartupPath + "\\ExcelTemplates\\SemesterTemplate.xltx";
				string semesterParam=Utilities.getSemesterConditionString(semester);
                float FIT_sum = 0.0f, MSF_sum = 0.0f, Mag_Sum = 0.0f, IDPO_sum = 0.0f;
                float FIT_lecSum = 0.0f, MSF_lecSum = 0.0f, Mag_lecSum = 0.0f, IDPO_lecSum = 0.0f;
                float contractSum = 0.0f, indPlanSum = 0.0f;

				//if (semester == "Осенний")
				//{
				//	semesterParam = "(Semester.Descr = '1'  OR  Semester.Descr = '3'  OR  Semester.Descr = '5'  OR  Semester.Descr = '7'  OR  Semester.Descr = '9'  OR  Semester.Descr = '11')";
				//}
				//else
				//{
				//	semesterParam = "(Semester.Descr = '2'  OR  Semester.Descr = '4'  OR  Semester.Descr = '6'  OR Semester.Descr = '8'  OR  Semester.Descr = '10'  OR  Semester.Descr = '12')";
				//}

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
							ObjWorkSheet.Cells[rowCounter, ApplicationSettings.ExportSettings.SemesterSetting.DisciplineCostColumn] = Math.Round( summ,2);
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
                    ObjWorkSheet.Cells[rowCounter, 5] = Math.Round(FIT_sum,2);
                    ObjWorkSheet.Cells[rowCounter, 6] = FIT_lecSum;

                    rowCounter = counts["ФИТ"] + counts["МСФ"] + counts["ИДПО"] + counts["МАГ"] + 33;
                    ObjWorkSheet.Cells[rowCounter, 5] = Math.Round(Mag_Sum,2);
                    ObjWorkSheet.Cells[rowCounter, 6] = Mag_lecSum;

                    rowCounter = counts["ФИТ"] + counts["МСФ"]+ counts["ИДПО"]+25;
                    ObjWorkSheet.Cells[rowCounter, 5] = Math.Round(IDPO_sum,2);
                    ObjWorkSheet.Cells[rowCounter, 6] = IDPO_lecSum;

                    rowCounter = counts["ФИТ"] + counts["МСФ"] + 17;
                    ObjWorkSheet.Cells[rowCounter, 5] = Math.Round(MSF_sum,2);
                    ObjWorkSheet.Cells[rowCounter, 6] = MSF_lecSum;

                   
                    rowCounter = counts["ФИТ"] + counts["ИДПО"] + counts["МАГ"] + counts["МСФ"] + 38;
                    ObjWorkSheet.Cells[rowCounter, 5] = Math.Round(indPlanSum,2);
                    ObjWorkSheet.Cells[rowCounter+1, 5] = Math.Round(contractSum,2);
                    ObjWorkSheet.Cells[rowCounter+2, 5] = Math.Round(indPlanSum +contractSum,2);

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
                //CafedralDB.SourceCode.Model.Exporter.ExportWorkload.Export(year);
                //return;

                string path = System.Windows.Forms.Application.StartupPath + "\\ExcelTemplates\\WorkloadTemplate.xltx";

				Excel.Application ObjExcel = new Excel.Application();
				Excel.Workbook ObjWorkBook;
				Excel.Worksheet ObjWorkSheet;
                
				//Книга.
				ObjWorkBook = ObjExcel.Workbooks.Add(path);//System.Reflection.Missing.Value);
														   //Таблица.
				ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
				ObjWorkBook.Title = string.Format("Нагрузка ({0} / {1})", year, Convert.ToInt32(year) + 1);
				//

				for (int i = 1; i < Math.Min(ObjWorkBook.Sheets.Count, 18); i++)
				{
					ObjWorkBook.Sheets[i].Cells[1, 17] = string.Format("{0} / {1}", year, Convert.ToInt32(year) + 1);
				}

				string strSQL = "SELECT StudyYear.StudyYear, Workload.ID, StudyForm.ID AS Expr2, Qualification.ID AS Expr1, " +
                    "Discipline.Descr, Speciality.Descr AS Expr4, Faculty.Descr AS Expr3, Semester.Descr AS Expr5, [Group].StudentCount, " +
                    "Semester.WeekCount, Discipline.LectureCount, Discipline.PracticeCount, Discipline.LabCount, Discipline.Ekz, Discipline.Zach, " +
                    "Discipline.RGR, Discipline.KR, Discipline.KP,  UchPr, prPr, predDipPr, KonsZaoch, GEK, GAK, GAKPred, DPruk, DopuskVKR, " +
                    "RetzVKR, DPretz, ASPRuk, MAGRuk, MAGRetz, RukKaf, NIIR, isSpecial,[Group].SubgroupCount,Contr " +
                    "FROM((((((((StudyYear INNER JOIN Workload ON StudyYear.ID = Workload.StudyYear) " +
                    "INNER JOIN Discipline ON Workload.Discipline = Discipline.ID) INNER JOIN [Group] ON Workload.[Group] = [Group].ID) " +
                    "INNER JOIN Qualification ON[Group].Qualification = Qualification.ID) INNER JOIN StudyForm ON[Group].StudyForm = StudyForm.ID) " +
                    "INNER JOIN Speciality ON[Group].Speciality = Speciality.ID) INNER JOIN Faculty ON Speciality.Faculty = Faculty.ID) " +
                    "INNER JOIN Semester ON Workload.Semester = Semester.ID) WHERE(StudyYear.StudyYear = @year) ORDER BY Workload.ID";

                /*
                 Reader:

                    0 - StudyYear                   1 - Workload.ID
                    2 - StudyForm.ID                3 - Qualification.ID
                    4 - Discipline.Descr            5 - Speciality.Descr
                    6 - Faculty.Descr               7 - Semester.Descr
                    8 - [Group].StudentCount        9 - Semester.WeekCount
                    10 - Discipline.LectureCount    11 - Discipline.PracticeCount
                    12 - Discipline.LabCount        13 - Discipline.Ekz
                    14 - Discipline.Zach            15 - Discipline.RGR
                    16 - Discipline.KR              17 - Discipline.KP
                    18 - UchPr                      19 - prPr
                    20 - predDipPr                  21 - KonsZaoch
                    22 - GEK                        23 - GAK
                    24 - GAKPred                    25 - DPruk
                    26 - DopuskVKR                  27 - RetzVKR
                    28 - DPretz                     29 - ASPRuk
                    30 - MAGRuk                     31 - MAGRetz
                    32 - RukKaf                     33 - NIIR
                    34 - isSpecial                  35 - [Group].SubgroupCount
                    36 - Contr
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

				Dictionary<Excel.Worksheet, int> rowCounters = new Dictionary<Excel.Worksheet, int>();
				//пишем данные
				for (int i = 2; i <= 17; i++)
				{
					rowCounters.Add(ObjWorkBook.Sheets[i], 6);
				}

				while (reader.Read())
				{
					int workloadID = Convert.ToInt32(reader[1]);
					//Определим листы для записи
					List<Excel.Worksheet> toWriteList = new List<Excel.Worksheet>();
                    if (reader[6].ToString() == "МСФ")
                    {
                        toWriteList.Add(ObjWorkBook.Sheets[6]);
                    }
                    else
                    {
                        if (Convert.ToInt32(reader[2]) == 2)
                        {
                            toWriteList.Add(ObjWorkBook.Sheets[5]);
                        }
                        else
                        {
                            if (Convert.ToInt32(reader[3]) == 1)
                            {
                                toWriteList.Add(ObjWorkBook.Sheets[4]);
                            }
                            else
                            {
                                toWriteList.Add(ObjWorkBook.Sheets[8]);
                            }
                        }
                    }

					Dictionary<Excel.Worksheet, WorkloadAssign> assignsForTeachers = new Dictionary<Excel.Worksheet, WorkloadAssign>(); 
               
					List<WorkloadAssign> assigns = DataManager.SharedDataManager().GetWorkloadAssigns(Convert.ToInt32(reader[1]));
					if (assigns.Count != 0)
					{
						foreach (WorkloadAssign assign in assigns)
						{
							string teacherName = DataManager.SharedDataManager().GetEmployee(assign.EmployeeID).Name;
                            Excel.Worksheet sheet = ObjWorkBook.Sheets[teacherName];
							if (sheet != null)
							{
								toWriteList.Add(sheet);
								assignsForTeachers.Add(sheet,assign);
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
                    bool flag = true;
					foreach (Excel.Worksheet sheet in toWriteList)
					{
						int countStud = Convert.ToInt32(reader[8]);
                        if (sheet.Index > 8 && Convert.ToBoolean(reader[34]))
                        {
                            WorkloadAssign assign = DataManager.SharedDataManager().
                                GetWorkloadAssign(Convert.ToInt32(reader[1]), sheet.Name);
                            countStud = assign.StudentsCount;

                        }
                        //TestWorkload(Convert.ToInt32(reader[1]));/*DEBUG*/
                        ExportNotAssign(reader, sheet, rowCounters,countStud);
					}
				}
                //object misValue = System.Reflection.Missing.Value;
                //var nameWorkbook = "Workload_";
                ObjExcel.Visible = true;
                ObjExcel.UserControl = true;

            /*
                ObjWorkBook.Close(true, nameWorkbook, misValue);
                ObjExcel.Quit();

                CafedralDB.SourceCode.Utilities.ReleaseObject(ObjWorkSheet);
                CafedralDB.SourceCode.Utilities.ReleaseObject(ObjWorkBook);
                CafedralDB.SourceCode.Utilities.ReleaseObject(ObjExcel);
            */
            }

            /*DEBUG*/
            static void TestWorkload(int workloadId)
            {
                var calc = Calculator.GetWorkloadCost(workloadId);
                string res = "";
                int distId = DataManager.SharedDataManager().GetWorkload(workloadId).Discipline;
                res += "WorkLoad " + " - " + DataManager.SharedDataManager().GetDiscipline(distId).Descr + "\r\n";
                res += "GEKCost " + " - " + calc.GEKCost + "\r\n";
                res += "GAKCost " + " - " + calc.GAKCost + "\r\n";
                res += "GosEkz " + " - " + calc.GosEkz + "\r\n\r\n";

                File.WriteAllText(System.AppDomain.CurrentDomain.BaseDirectory + "\\testWorkload\\" + calc.WorkLoadID + ".txt", res);
            }

			private static void ExportNotAssign(OleDbDataReader reader, Excel.Worksheet sheet, Dictionary<Excel.Worksheet, int> rowCounters, int countStud)
			{
                CafedralDB.SourceCode.Settings.CalculationSetting calculationSetting = new CafedralDB.SourceCode.Settings.CalculationSetting();
                var workloadCost = Calculator.GetWorkloadCost(Convert.ToInt32(reader[1]));
                int groupCount = Convert.ToInt32(reader[35]);
                int studentCount= Convert.ToInt32(reader[8]);
                int semestr = Convert.ToInt32(reader[7].ToString());
                var subGroupCount= groupCount == 1 ? Math.Floor((double)studentCount / 9)>0? Math.Floor((double)studentCount / 9):1 :
                    Math.Floor((double)studentCount / groupCount / 9) * groupCount;
                sheet.Cells[rowCounters[sheet], 1] = rowCounters[sheet] - 5;
				sheet.Cells[rowCounters[sheet], 2] = reader[4].ToString();//Дисциплина
				sheet.Cells[rowCounters[sheet], 3] = reader[5].ToString();//Специальность
				sheet.Cells[rowCounters[sheet], 4] = reader[6].ToString();//Факультет
                sheet.Cells[rowCounters[sheet], 5] = reader[2].ToString() == "1" ? "о" : "з";//Форма обучения
                sheet.Cells[rowCounters[sheet], 6] = semestr > 8 ? semestr - 8 : semestr;//reader[7].ToString();//Семестр
				sheet.Cells[rowCounters[sheet], 7] = countStud;//reader[8].ToString(); Кол-во студентов
				sheet.Cells[rowCounters[sheet], 8] = reader[9].ToString(); // Кол-во недель
                sheet.Cells[rowCounters[sheet], 9] = reader[35].ToString();//reader[10].ToString(); // Кол-во групп
                sheet.Cells[rowCounters[sheet], 10] = Convert.ToInt32(subGroupCount).ToString();// Кол-во подгрупп
                //sheet.Cells[rowCounters[sheet], 10] = Math.Floor(Convert.ToDouble(reader[8]) / 9).ToString(); 
                sheet.Cells[rowCounters[sheet], 11] = reader[10].ToString(); // Кол-во лекций
				sheet.Cells[rowCounters[sheet], 12] = reader[11].ToString(); // Кол-во практ. работ
				sheet.Cells[rowCounters[sheet], 13] = reader[12].ToString(); // Кол-во лаб. работ
				sheet.Cells[rowCounters[sheet], 14] = Convert.ToBoolean(reader[13]) ? "+" : "";//Экзамен
				sheet.Cells[rowCounters[sheet], 15] = Convert.ToBoolean(reader[14]) ? "+" : "";// Зачет
				sheet.Cells[rowCounters[sheet], 16] = Convert.ToBoolean(reader[15]) ? "+" : "";// РГр
				sheet.Cells[rowCounters[sheet], 17] = Convert.ToBoolean(reader[16]) ? "+" : "";// КР
				sheet.Cells[rowCounters[sheet], 18] = Convert.ToBoolean(reader[17]) ? "+" : "";// КП

                sheet.Cells[rowCounters[sheet], 19] = Convert.ToBoolean(reader[18]) ? "+" : "";// Учебная практика
                sheet.Cells[rowCounters[sheet], 20] = Convert.ToBoolean(reader[19]) ? "+" : "";// Производственная практика
                sheet.Cells[rowCounters[sheet], 21] = Convert.ToBoolean(reader[20]) ? "+" : "";// Преддипломная практика
                sheet.Cells[rowCounters[sheet], 23] = Convert.ToBoolean(reader[22]) ? "+" : "";// ГЭК
                sheet.Cells[rowCounters[sheet], 24] = Convert.ToBoolean(reader[23]) ? "+" : "";// Гак
                sheet.Cells[rowCounters[sheet], 25] = Convert.ToBoolean(reader[24]) ? "+" : "";// Гак председатель
                sheet.Cells[rowCounters[sheet], 33] = Convert.ToBoolean(reader[32]) ? "+" : "";// Руководство кафедрой







                /*
                    Заполнение часов начинается с ячейки 34 и имеет след порядок:
                        - лекции -34
                        - практ. работы - 35
                        - лаб.работы - 36
                        - Консультации - 37
                        - экзамен - 38
                        - Зачет - 39
                        - РГР - 40
                        - КР - 41
                        - КП - 42
                        - Уч.практика - 43
                        - Произв. практика - 44
                        - Преддипл.практика - 45
                        - ГЭК - 47
                        - ГАК - 48
                        - ВКР бакалавры - 50
                        - Магистерская диссертация - 56

                */

                sheet.Cells[rowCounters[sheet], 34] = Math.Round(workloadCost.LectureCost,2); //часов лекций
                sheet.Cells[rowCounters[sheet], 35] = Math.Round(workloadCost.PracticeCost, 2); //часов практ. работ
                sheet.Cells[rowCounters[sheet], 36] = Math.Round(workloadCost.LabCost, 2); // часов лабораторных
                sheet.Cells[rowCounters[sheet], 37] = Math.Round(workloadCost.KonsCost, 2); //часов консультаций
                sheet.Cells[rowCounters[sheet], 38] = Math.Round(workloadCost.EkzCost, 2); // часов экзамен
                sheet.Cells[rowCounters[sheet], 39] = Math.Round(workloadCost.ZachCost, 2); //часов зачет
                sheet.Cells[rowCounters[sheet], 40] = 0;
                sheet.Cells[rowCounters[sheet], 41] = Math.Round(workloadCost.KRCost, 2); //часов курсовая работа
                sheet.Cells[rowCounters[sheet], 42] = Math.Round(workloadCost.KPCost, 2); //часов курсовой проект

                //Учебная практика
                if (workloadCost.UchPracCost != 0)
                {
                    sheet.Cells[rowCounters[sheet], 8] = reader[18].ToString(); // Кол-во недель
                    sheet.Cells[rowCounters[sheet], 19] = Convert.ToBoolean(reader[18]) ? "+" : ""; 
                    sheet.Cells[rowCounters[sheet], 43] = Math.Round(workloadCost.UchPracCost, 2);
                }
                //Производственная практика
                if (workloadCost.PrPracCost != 0)
                {
                    sheet.Cells[rowCounters[sheet], 8] = reader[19].ToString(); // Кол-во недель
                    sheet.Cells[rowCounters[sheet], 20] = Convert.ToBoolean(reader[19]) ? "+" : "";
                    sheet.Cells[rowCounters[sheet], 44] = Math.Round(workloadCost.PrPracCost, 2);
                }
                //Преддипломная практика
                if (workloadCost.PredDipPracCost != 0)
                {
                    sheet.Cells[rowCounters[sheet], 8] = reader[20].ToString(); // Кол-во недель
                    sheet.Cells[rowCounters[sheet], 21] = Convert.ToBoolean(reader[20]) ? "+" : "";
                    sheet.Cells[rowCounters[sheet], 45] = Math.Round(workloadCost.PredDipPracCost, 2);
                }
                //Научно-исследовательская работа
                if (workloadCost.NIIRCost != 0)
                {
                    sheet.Cells[rowCounters[sheet], 8] = reader[33].ToString(); // Кол-во недель
                    sheet.Cells[rowCounters[sheet], 22] = Convert.ToBoolean(reader[33]) ? "+" : "";
                    sheet.Cells[rowCounters[sheet], 45] = Math.Round(workloadCost.NIIRCost, 2);
                }
                //ГЕК
                if (workloadCost.GEKCost != 0)
                {
                    sheet.Cells[rowCounters[sheet], 47] = Math.Round(workloadCost.GEKCost, 2);
                }
                //гос Экзамен
                if (workloadCost.GosEkz != 0)
                {
                    sheet.Cells[rowCounters[sheet], 47] = Math.Round(workloadCost.GosEkz, 2);

                }
                //ГАК
                if (workloadCost.GAKCost != 0)
                {
                    sheet.Cells[rowCounters[sheet], 48] = Math.Round(workloadCost.GAKCost, 2);
                }
                //ГAКпред
                if (workloadCost.GAKPredCost != 0)
                {
                    sheet.Cells[rowCounters[sheet], 49] = Math.Round(workloadCost.GAKPredCost, 2);
                }
                // Дипломная работа 
                if (workloadCost.DPRukCost != 0)
                {
                    sheet.Cells[rowCounters[sheet], 50]= Math.Round(workloadCost.DPRukCost, 2);
                }
                //Диссертация магистров
                if (workloadCost.MagRuk != 0)
                {
                    sheet.Cells[rowCounters[sheet], 55] = Math.Round(workloadCost.MagRuk, 2);
                }
                //Руководство кафедрой
                if (workloadCost.RukKafCost != 0)
                {
                    sheet.Cells[rowCounters[sheet], 57] = Math.Round(workloadCost.RukKafCost, 2);
                }


                sheet.Cells[rowCounters[sheet], 51] = Convert.ToBoolean(reader[26]) ?
                    (4 * countStud).ToString() : "";//рец дисс.

				sheet.Cells[rowCounters[sheet], 54] = Convert.ToBoolean(reader[29]) ?
                    (countStud * calculationSetting.AspRuk) : 0f; //AspRuk

                string disciplineName = reader[4].ToString().ToLower();

				//sheet.Cells[rowCounters[sheet], 49] = Convert.ToBoolean(reader[24]) ? 
    //                (ApplicationSettings.CalculationSettings.GAKPred * countStud) : 0;//ГAКпред.
				//sheet.Cells[rowCounters[sheet], 50] = Convert.ToBoolean(reader[25]) ?
    //                (ApplicationSettings.CalculationSettings.DPruk * countStud) : 0;//Дипл проект руководство
				/*
				if (disciplineName.ToLower().Contains("норм") && disciplineName.ToLower().Contains("маг"))
					sheet.Cells[rowCounters[sheet], 55] = countStud * ApplicationSettings.CalculationSettings.NormocontrolMag;

				if (disciplineName.ToLower().Contains("доп") && disciplineName.ToLower().Contains("маг"))
					sheet.Cells[rowCounters[sheet], 55] = countStud * ApplicationSettings.CalculationSettings.DopuskDissMag;
                */
				//sheet.Cells[rowCounters[sheet], 55] = Convert.ToBoolean(reader[30]) ?
                //(30 * Convert.ToInt32(reader[30])).ToString() : "";//рук маг

                /*
				if (disciplineName.ToLower().Contains("норм") && disciplineName.ToLower().Contains("маг"))
					sheet.Cells[rowCounters[sheet], 55] = countStud * ApplicationSettings.CalculationSettings.NormocontrolMag;

				if (disciplineName.ToLower().Contains("доп") && disciplineName.ToLower().Contains("маг"))
					sheet.Cells[rowCounters[sheet], 55] = countStud * ApplicationSettings.CalculationSettings.DopuskDissMag;
                */
				rowCounters[sheet]++;
			}

			private static void ExportAssigned(OleDbDataReader reader, Excel.Worksheet sheet, Dictionary<Excel.Worksheet, int> rowCounters, int countStud)
			{

                CafedralDB.SourceCode.Settings.CalculationSetting calculationSetting = new CafedralDB.SourceCode.Settings.CalculationSetting();
				string disciplineName = reader[4].ToString();
				sheet.Cells[rowCounters[sheet], 1] = rowCounters[sheet] - 5;
				sheet.Cells[rowCounters[sheet], 2] = reader[4].ToString();
				sheet.Cells[rowCounters[sheet], 3] = reader[5].ToString();
				sheet.Cells[rowCounters[sheet], 5] = reader[2].ToString() == "1" ? "о" : "з";
				sheet.Cells[rowCounters[sheet], 4] = reader[6].ToString();
				sheet.Cells[rowCounters[sheet], 6] = reader[7].ToString();
				sheet.Cells[rowCounters[sheet], 7] = countStud;//reader[8].ToString();
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
				sheet.Cells[rowCounters[sheet], 19 + 23] = Convert.ToInt32(reader[18]) != 0 ? (Convert.ToInt32(reader[18]) * calculationSetting.UchPr) : 0;//уч пр.
				sheet.Cells[rowCounters[sheet], 20 + 23] = Convert.ToInt32(reader[19]) != 0 ? (Convert.ToInt32(reader[19]) * calculationSetting.PrPr) : 0;//пр пр.
				sheet.Cells[rowCounters[sheet], 21 + 23] = Convert.ToInt32(reader[20]) != 0 ? (Convert.ToInt32(reader[20]) * calculationSetting.PreddipPr) : 0;//преддип пр.
				sheet.Cells[rowCounters[sheet], 22 + 23] = Convert.ToInt32(reader[33]) != 0 ? (Convert.ToInt32(reader[33]) * calculationSetting.NIIR * countStud) : 0;//нир
				((Excel.Range)(sheet.Cells[rowCounters[sheet], 24 + 23])).Value = Convert.ToBoolean(reader[22]) ? (calculationSetting.GEK * countStud) : 0;//ГЭК.
				sheet.Cells[rowCounters[sheet], 25 + 23] = Convert.ToBoolean(reader[23]) ? (calculationSetting.GAK * countStud) : 0;//ГAК.
				sheet.Cells[rowCounters[sheet], 26 + 23] = Convert.ToBoolean(reader[24]) ? (calculationSetting.GAKPred * countStud) : 0;//ГAКпред.
				sheet.Cells[rowCounters[sheet], 27 + 23] = Convert.ToBoolean(reader[25]) ? (calculationSetting.DPruk * countStud) : 0;//допуск
				sheet.Cells[rowCounters[sheet], 28 + 23] = Convert.ToBoolean(reader[26]) ? (4 * countStud).ToString() : "";//рец дисс.
				sheet.Cells[rowCounters[sheet], 32 + 23] = Convert.ToBoolean(reader[30]) ? (30 * Convert.ToInt32(reader[30])).ToString() : "";//рук маг
				sheet.Cells[rowCounters[sheet], 31 + 23] = Convert.ToBoolean(reader[29]) ? (countStud * calculationSetting.AspRuk) : 0f;
				if (disciplineName.ToLower().Contains("норм") && disciplineName.ToLower().Contains("маг"))
					sheet.Cells[rowCounters[sheet], 32 + 23] = countStud * calculationSetting.NormocontrolMag;
				if (disciplineName.ToLower().Contains("доп") && disciplineName.ToLower().Contains("маг"))
					sheet.Cells[rowCounters[sheet], 32 + 23] = countStud * calculationSetting.DopuskDissMag;
				if (disciplineName.ToLower().Contains("рук") && (disciplineName.ToLower().Contains("нир")|| disciplineName.ToLower().Contains("ниир")))
					sheet.Cells[rowCounters[sheet], 22 + 23] = Convert.ToInt32(reader[33]) != 0 ? (Convert.ToInt32(reader[33]) * calculationSetting.NIIRRukAsp * countStud) : 0;//нир
				rowCounters[sheet]++;
			}
		}
        [Obsolete("Старый класс экспорта контракта")]
        public static class Contract
        {
            public static void ExportContract(int teacherID, string year, string semester)
            {
                string semesterParam="";
                if (semester == "Осенний")
                {
                    semesterParam = "(Semester.Descr = '1'  OR  Semester.Descr = '3'  OR  Semester.Descr = '5'  OR  Semester.Descr = '7'  OR  Semester.Descr = '9'  OR  Semester.Descr = '11')";
                }
                else
                {
                    semesterParam = "(Semester.Descr = '2'  OR  Semester.Descr = '4'  OR  Semester.Descr = '6'  OR Semester.Descr = '8'  OR  Semester.Descr = '10'  OR  Semester.Descr = '12')";
                }

                string strSQL = "SELECT Discipline.Descr, Faculty.Descr AS Expr1, Speciality.Descr AS Expr2,[Group].Descr, [Group].EntryYear, " +
                    "[Group].StudentCount, Semester.WeekCount, DisciplineType.Descr AS Expr3, StudyYear.StudyYear, WorkloadAssign.Teacher, " +
                    "Discipline.KonsZaoch, Discipline.GAKPred, Discipline.DPruk, Discipline.DopuskVKR, Discipline.RetzVKR, Discipline.DPretz," +
                    " Discipline.ASPRuk, Discipline.MAGRuk, Discipline.MAGRetz, Discipline.RukKaf, Discipline.NIIR, Discipline.isSpecial, " +
                    "Workload.ID FROM(StudyYear INNER JOIN ((Faculty INNER JOIN Speciality ON Faculty.ID = Speciality.Faculty)" +
                    " INNER JOIN(Semester INNER JOIN (Qualification INNER JOIN ((DisciplineType INNER JOIN Discipline ON DisciplineType.ID = Discipline.DisciplineType)" +
                    " INNER JOIN(Workload INNER JOIN[Group] ON Workload.[Group] = Group.ID) ON Discipline.ID = Workload.Discipline)" +
                    " ON Qualification.ID = Group.Qualification) ON Semester.ID = Workload.Semester) ON Speciality.ID = Group.Speciality)" +
                    " ON StudyYear.ID = Workload.StudyYear) INNER JOIN WorkloadAssign ON Workload.ID = WorkloadAssign.Workload" +
                    " WHERE(Discipline.Contr = TRUE) AND(StudyYear.StudyYear=[@param1]) AND(WorkloadAssign.Teacher=[@param2])";// AND " + semesterParam;
                DataManager.SharedDataManager();
                var cn = new System.Data.OleDb.OleDbConnection(DataManager.Connection.ConnectionString);
                var cmd = new System.Data.OleDb.OleDbCommand();
                cn.Open();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSQL;
                cmd.Parameters.AddWithValue("@param1", year);
                cmd.Parameters.AddWithValue("@param2", teacherID);
                System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
                string teacherName=DataManager.SharedDataManager().GetEmployee(teacherID).Name;

                List<ContractDiscr> DisciplineList = new List<ContractDiscr>();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        WorkloadCost workloadCost = Calculator.GetWorkloadCost(Convert.ToInt32(reader[22]));
                        ContractDiscr contractDescr = new ContractDiscr(workloadCost);
                        contractDescr.Group = reader[3].ToString();
                        contractDescr.Discipline = reader[0].ToString();
                        DisciplineList.Add(contractDescr);
                    }
                }
                PrintContract(teacherName, DisciplineList, year);
                
                //MessageBox.Show("Документ создан.");
            }

            static void PrintContract(string TeacherName, List<ContractDiscr> DisciplineList, string year)
            {
                string path = System.Windows.Forms.Application.StartupPath + "\\WordTemplate\\ContractTemlate.docx";
                TimeSpan nowTime = DateTime.Now.TimeOfDay;
                string nowDate = DateTime.Now.Date.ToShortDateString();
                string dateWithTime = string.Format("{0}-{1}",nowDate,nowTime.ToString(@"hh\.mm\.ss"));
                string filename = string.Format("Contract of {0} - {1}({2}).docx", TeacherName, year, dateWithTime);
                string outpath = System.Windows.Forms.Application.StartupPath + "\\Contracts\\" + filename;
                string[] prefixTags = ApplicationSettings.ExportSettings.ContractSetting.GetPrefixTag();
                File.Copy(path, outpath);               
                TableContent tableContent = new TableContent("Discipline table");
                for(int i=0;i<7;i++)
                {
                    tableContent.AddRow(
                        new FieldContent("row number", (i+1).ToString()),
                        new FieldContent("Discipline", FindItem(DisciplineList, i, "Discipline")),
                        new FieldContent("Group", FindItem(DisciplineList, i, "Group"))
                        );
                }
                List<FieldContent> fieldContents = FillTableOfTrainingTask(DisciplineList, prefixTags);
                var valueToFill = new Content();

                valueToFill.Fields.Add(new FieldContent("Teacher name", TeacherName));
                valueToFill.Tables.Add(tableContent);

                foreach (FieldContent fieldContent in fieldContents)
                {
                    valueToFill.Fields.Add(fieldContent);
                }

                using (var outputDocument = new TemplateProcessor(outpath)
                    .SetRemoveContentControls(true))
                {
                    outputDocument.FillContent(valueToFill);
                    outputDocument.SaveChanges();
                }
                
                var _path = System.Windows.Forms.Application.StartupPath+"\\" + outpath;
                System.Diagnostics.Process.Start(outpath);
            }

            static List<FieldContent> FillTableOfTrainingTask(List<ContractDiscr> contractDiscrs,string[] prefixTags)
            {
                List<FieldContent> fieldContents = new List<FieldContent>();
                double totalAll = 0, totalColumn = 0;
                double value;
                string tempValue ="";
                double[] totalRow = new double[prefixTags.Length];
                string tag = "";
                for(int i=0;i<10;i++)
                {
                    for (int j = 0; j < prefixTags.Length;j++)
                    {
                        tag = prefixTags[j] + "_" + i;
                        tempValue = FindItem(contractDiscrs, i, prefixTags[j]);
                        value = tempValue == "" ? 0 : Convert.ToDouble(tempValue);
                        totalColumn += value;
                        totalRow[j] += value;
                        if (prefixTags[j] == "Total")
                            value = totalColumn;
                        fieldContents.Add(new FieldContent(tag, value == 0 ? "" : value.ToString()));
                    }
                    totalColumn = 0;
                }
                for (int j = 0; j < prefixTags.Length; j++)
                {
                    if (prefixTags[j] != "Total")
                    {
                        tag = prefixTags[j] + "_total";
                        totalAll += totalRow[j];
                        fieldContents.Add(new FieldContent(tag, totalRow[j] == 0 ? "" : totalRow[j].ToString()));
                    }
                }

                tag = ApplicationSettings.ExportSettings.ContractSetting.TotalTimeTag;
                fieldContents.Add(new FieldContent(tag,totalAll.ToString()));

                return fieldContents;
            }

            static string FindItem(List<ContractDiscr> list,int index,string field)
            {
                switch(field)
                {
                    case "Discipline":
                        if (index >= list.Count)
                            return "";
                        else
                            return list[index].Discipline;
                    case "Group":
                        if (index >= list.Count)
                            return "";
                        else
                            return list[index].Group;
                    case "Lecture":
                        return index >= list.Count ? "" : list[index].LectureCost.ToString();
                    case "Pract":
                        return index >= list.Count ? "" : list[index].PracticeCost.ToString();
                    case "Lab":
                        return index >= list.Count ? "" : list[index].LabCost.ToString();
                    case "Kr":
                        return index >= list.Count ? "" : list[index].KRCost.ToString();
                    case "Kp":
                        return index >= list.Count ? "" : list[index].KPCost.ToString();
                    case "Nir":
                        return index >= list.Count ? "" : list[index].NIIRCost.ToString();
                    case "RukPract":
                        return index >= list.Count ? "" : (list[index].PredDipPracCost 
                            + list[index].PrPracCost + list[index].UchPracCost).ToString();
                    case "Cons":
                        return index >= list.Count ? "" : list[index].KonsCost.ToString();
                    case "Zach":
                        return index >= list.Count ? "" : list[index].ZachCost.ToString();
                    case "Ekz":
                        return index >= list.Count ? "" : list[index].EkzCost.ToString();
                    case "Vkr":
                        return index >= list.Count ? "" : list[index].DPRukCost.ToString();
                    case "Gek":
                        return index >= list.Count ? "" : list[index].GEKCost.ToString();
                    case "RecVkr":
                        return index >= list.Count ? "" : "";
                    default:
                        return "";
                }
            }
        }
	}
}

