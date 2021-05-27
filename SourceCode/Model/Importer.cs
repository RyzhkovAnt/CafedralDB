﻿using Model;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ApplicationSettings;
using Model.Entity;
namespace CafedralDB.SourceCode.Model
{
    static class Importer
    {
        public static void ImportDataFromExcel(string path, string year)
        {
            List<Discipline> disciplines = new List<Discipline>();
			List<Workload> workloads = new List<Workload>();
            #region Read disciplines from Excel
            //opening Excel
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

			//Reading data
			bool res = true;
			HashSet<string> answer =new HashSet<string>();
			//Проход по обоим листам
			for (int i = 1; i <= 1; i++)
            {
                xlWorkSheet = xlWorkBook.Worksheets[i];

                int counter = ImportSettings.StartReadingRow;

				while (xlWorkSheet.GetCellText(counter, 8) != "")
                {
                    string group = xlWorkSheet.GetCellText(counter, ImportSettings.GroupColumn);

                    int groupCount = Convert.ToInt32(xlWorkSheet.GetCellText(counter, ImportSettings.GroupCountColumn));
                    int studentCount = Convert.ToInt32(xlWorkSheet.GetCellText(counter, ImportSettings.StudentCountColumn));

                    string semester = xlWorkSheet.GetCellText(counter, ImportSettings.SemesterColumn);
                    int weeks = xlWorkSheet.GetCellText(counter, ImportSettings.WeeksColumn) != "" ? Convert.ToInt32(xlWorkSheet.GetCellText(counter, ImportSettings.WeeksColumn)) : 0;
                    string disciplineName = xlWorkSheet.GetCellText(counter, ImportSettings.DisciplineNameColumn);
                    int lectures = xlWorkSheet.GetCellText(counter, ImportSettings.LecturesColumn)!=""?Convert.ToInt32(xlWorkSheet.GetCellText(counter, ImportSettings.LecturesColumn)):0;
                    int labs = xlWorkSheet.GetCellText(counter, ImportSettings.LabsColumn)!=""? Convert.ToInt32(xlWorkSheet.GetCellText(counter, ImportSettings.LabsColumn)):0;
                    int practices = xlWorkSheet.GetCellText(counter, ImportSettings.PracticesColumn)!=""? Convert.ToInt32(xlWorkSheet.GetCellText(counter, ImportSettings.PracticesColumn)):0;
                    bool kz = xlWorkSheet.GetCellText(counter, ImportSettings.KzColumn)!="";
                    bool kr = xlWorkSheet.GetCellText(counter, ImportSettings.KrColumn)!="";
                    bool kp = xlWorkSheet.GetCellText(counter, ImportSettings.KpColumn)!="";
                    bool ekz = xlWorkSheet.GetCellText(counter, ImportSettings.EkzColumn)!="";
                    bool zach = xlWorkSheet.GetCellText(counter, ImportSettings.ZachColumn)!="";
                    //Проверить
					bool isSpecial = xlWorkSheet.GetCellText(counter, ImportSettings.ZachColumn+2) == "";

                    bool isContract = xlWorkSheet.GetCellText(counter, ImportSettings.ContractColumn) != "";

					Discipline discipline = new Discipline(counter - ImportSettings.StartReadingRow);
					int disciplineType = 1;
					
					discipline.Descr = disciplineName;
					if (!isSpecial)
					{
						discipline.LectureCount = lectures;
						discipline.PracticeCount = practices;
						discipline.LabCount = labs;
						discipline.KR = kr;
						discipline.KP = kp;
						discipline.Ekz = ekz;
						discipline.Zach = zach;
						discipline.Kz = kz;
                        discipline.Contract = isContract;
					}
					else
					{
						discipline.Special = true;

                        if (disciplineName.ToLower().Contains("практ"))
                        {
                            //if (disciplineName.ToLower().Contains("нир") || disciplineName.ToLower().Contains("исслед"))
                            //{
                            //    discipline.NIIR = practices;
                            //    disciplineType = 3;
                            //}
                            //else
							if (disciplineName.ToLower().Contains("уч"))
                            {
                                discipline.UchPr = practices;
                                disciplineType = 3;
                            }
                            else if (disciplineName.ToLower().Contains("преддип"))
                            {
                                discipline.PredDipPr = practices;
                                disciplineType = 3;
                            }
                            else if (disciplineName.ToLower().Contains("произв"))
                            {
                                discipline.PrPr = practices;
                                disciplineType = 3;
                            }
                        }

						if (disciplineName.ToLower().Contains("конс") && disciplineName.ToLower().Contains("заоч"))
						{
							discipline.KonsZaoch = true;
							disciplineType = 3;
						}

						if (disciplineName.ToLower().Contains("вып") && disciplineName.ToLower().Contains("раб"))
						{
							discipline.DPRuk = true;
							disciplineType = 2;
						}

						//if (disciplineName.ToLower().Contains("гэк") || (disciplineName.ToLower().Contains("гос") && disciplineName.ToLower().Contains("экз")))
						//{
						//	discipline.GEK = true;
						//	disciplineType = 3;
						//}
						if(disciplineName.ToLower().Contains("гос") && disciplineName.ToLower().Contains("экз"))
                        {
							disciplineType = 3;
							//Добавить в таблицу и классы Гос Экзамен
                        }
						if (disciplineName.ToLower().Contains("гак") )
						{
							if (disciplineName.ToLower().Contains("предс"))
							{
								discipline.GAKPred = true;
							}
							else
							{
								discipline.GAK = true; 
							}
							disciplineType = 3;
						}

                        if (disciplineName.ToLower().Contains("диссер"))
                        {
                            discipline.DPRuk = true;
                        }
                        if ((disciplineName.ToLower().Contains("науч") && disciplineName.ToLower().Contains("иссл"))|| disciplineName.ToLower().Contains("нир"))
						{
							discipline.NIIR = practices;
							disciplineType = 2;
						}

						if (disciplineName.ToLower().Contains("рук") && disciplineName.ToLower().Contains("каф"))
						{
							discipline.RukKaf = true;
							disciplineType = 3;
						}

						if (disciplineName.ToLower().Contains("рук") && disciplineName.ToLower().Contains("асп"))
						{
							discipline.ASPRuk = true;
							disciplineType = 2;
						}

						if ((disciplineName.ToLower().Contains("рук") || (disciplineName.ToLower().Contains("дисс"))) && disciplineName.ToLower().Contains("маг"))
						{
							discipline.MAGRuk = true;
							disciplineType = 2;
						}
					}
					discipline.DepartmentID = 1;

					discipline.TypeID = DataManager.SharedDataManager().FindTypeByName(disciplineName);

					//year - надо чтоб: int(year) - semester/2
					int course = 1;
                    //Ввести проверку, что если семестр болльше 8(для рассчета курса магистров)
                    //добавить Изменение длительности семестра
					if (semester!="")
						course = (int)Math.Ceiling(Convert.ToInt32(semester)/2f);
                    if (course >= 5)
                        course -= 4;
					int entryYear = (Convert.ToInt32(year) - course + 1);
					int specialityID = DataManager.SharedDataManager().GetSpecialityIDByName(group);
					int semesterID = DataManager.SharedDataManager().GetSemesterIDByName(semester);
					int yearID = DataManager.SharedDataManager().GetYearIDByName(year.ToString());
					int entryYearID = DataManager.SharedDataManager().GetYearIDByName(entryYear.ToString());
					int groupID = DataManager.SharedDataManager().GetGroupIDByYearAndSpeciality(entryYearID, specialityID);

					if (specialityID == -1)
					{
						answer.Add("Не найдена специальность - " + group + "\n");
						res = false;
					}

					if (semesterID == -1)
					{
						semesterID = 1;
						//answer.Add("Не найден семестр - " + semester + "\n");
						//res = false;
					}

					if (yearID == -1)
					{
						answer.Add("Не найден учебный год - " + year + "\n");
						res = false;
					}

					if (groupID == -1)
					{
						answer.Add("Не найдена группа - " + group + " поступившая в " + entryYear + "\n");
						res = false;
					}

					Workload workload = new Workload();
					workload.Semester = semesterID;
					workload.Year = yearID;
					workload.Group = groupID;
					
					if (res)
					{
						disciplines.Add(discipline);
						workloads.Add(workload);
					}
					res = true;
					counter++;

                    if(DataManager.SharedDataManager().GetSemester(semesterID).WeekCount!=weeks)
                    {
                        var newSemester = new Semester(semesterID);
                        newSemester.WeekCount = weeks;
                        newSemester.Name = semesterID.ToString();
                        DataManager.SharedDataManager().SetSemester(newSemester);
                    }

                    if(groupID!=-1 && (DataManager.SharedDataManager().GetGroup(groupID).SubgroupCount != groupCount ||
                        DataManager.SharedDataManager().GetGroup(groupID).StudentCount != studentCount))
                    {
                        var newGroup = new Group(groupID);
                        newGroup.SubgroupCount = groupCount;
                        newGroup.StudentCount = studentCount;
                        DataManager.SharedDataManager().SetGroup(newGroup);
                    }
				}
				
			}
			if (answer.Count > 0)
			{
				string log = "";
				foreach (string logstring in answer)
					log += logstring;
				MessageBox.Show(log);
			}
			if (DataManager.SharedDataManager().CheckDisciplinesAtYearExist(year))
			{
				DialogResult result=MessageBox.Show("Этот год уже присутствует!\nОчистить?", "Очистка данных", MessageBoxButtons.YesNo);
				if(result == DialogResult.Yes)
				{
					DataManager.SharedDataManager().ClearYear(year);
				}
			}

			for (int i=0;i<disciplines.Count;i++)
			{
				Discipline discipline = disciplines[i];
				Workload workload = workloads[i];

				int DisciplineID = DataManager.SharedDataManager().AddDiscipline(discipline);

				workload.Discipline = DisciplineID;

				int workloadID = DataManager.SharedDataManager().AddWorkload(workload);
				int specID = DataManager.SharedDataManager().GetSpecialityIDByGroupID(workload.Group);
				int lastYearAssign = DataManager.SharedDataManager().FindLastWorkloadAssign(workload, discipline.Descr, specID);

				if(lastYearAssign!=-1)
				{
					WorkloadAssign assign = new WorkloadAssign(0);
					assign.EmployeeID = lastYearAssign;
					assign.WorkloadID = workloadID;
					DataManager.SharedDataManager().AddWorkloadAssign(assign);

				}
			}

            //closing Excel
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Utilities.ReleaseObject(xlWorkSheet);// releaseObject(xlWorkSheet);
            Utilities.ReleaseObject(xlWorkBook);//releaseObject(xlWorkBook);
            Utilities.ReleaseObject(xlApp);//releaseObject(xlApp);
            #endregion

            //int DisciplineID = DataManager.SharedDataManager().AddDiscipline();
            //'запишем данные в таблицу Дисциплина
            //strSQL = "INSERT INTO Disciplina (descr,chair_id,lecture,practice,lab,KZzaoch,kr,kp,ekz,zach) VALUES (" & _
            //"'" & name & "'" & "," & "1," & lec & "," & pr & "," & lab & "," & kz & "," & kr & "," & kp & "," & ekz & ",""" & zach & """)"


            #region Past workload to Database

            #endregion
        }

		private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
