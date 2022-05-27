using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ApplicationSettings;
using CafedralDB.SourceCode.Model.Entity;
using CafedralDB.SourceCode.Settings;
namespace CafedralDB.SourceCode.Model
{
    static class Importer
    {
        public static void ImportDataFromExcel(string path, string year)
        {
            List<Discipline> disciplines = new List<Discipline>();
			List<Workload> workloads = new List<Workload>();

			ImportSetting importSettings = new ImportSetting();
			
            #region Read disciplines from Excel
            //opening Excel
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows,
				"\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

			//Reading data
			bool res = true;
			HashSet<string> answer =new HashSet<string>();
			//Проход по обоим листам
			for (int i = 1; i <= 1; i++)
            {
                xlWorkSheet = xlWorkBook.Worksheets[i];

                int counter = importSettings.StartReadingRow;

				while (xlWorkSheet.GetCellText(counter, 8) != "")
                {
                    string group = xlWorkSheet.GetCellText(counter, importSettings.GroupColumn);

                    int groupCount = Convert.ToInt32(xlWorkSheet.GetCellText(counter, importSettings.GroupCountColumn));
                    int studentCount = Convert.ToInt32(xlWorkSheet.GetCellText(counter, importSettings.StudentCountColumn));

                    string semester = xlWorkSheet.GetCellText(counter, importSettings.SemesterColumn);
                    int weeks = xlWorkSheet.GetCellText(counter, importSettings.WeeksColumn) != "" ? 
						Convert.ToInt32(xlWorkSheet.GetCellText(counter, importSettings.WeeksColumn)) : 0;
                    string disciplineName = xlWorkSheet.GetCellText(counter, importSettings.DisciplineNameColumn);
                    int lectures = xlWorkSheet.GetCellText(counter, importSettings.LecturesColumn)!=""?
						Convert.ToInt32(xlWorkSheet.GetCellText(counter, importSettings.LecturesColumn)):0;
                    int labs = xlWorkSheet.GetCellText(counter, importSettings.LabsColumn)!=""? 
						Convert.ToInt32(xlWorkSheet.GetCellText(counter, importSettings.LabsColumn)):0;
                    int practices = xlWorkSheet.GetCellText(counter, importSettings.PracticesColumn)!=""?
						Convert.ToInt32(xlWorkSheet.GetCellText(counter, importSettings.PracticesColumn)):0;
                    bool kz = xlWorkSheet.GetCellText(counter, importSettings.KzColumn)!="";
                    bool kr = xlWorkSheet.GetCellText(counter, importSettings.KrColumn)!="";
                    bool kp = xlWorkSheet.GetCellText(counter, importSettings.KpColumn)!="";
                    bool ekz = xlWorkSheet.GetCellText(counter, importSettings.EkzColumn)!="";
                    bool zach = xlWorkSheet.GetCellText(counter, importSettings.ZachColumn)!="";
                    
					//Проверить
					bool isSpecial = xlWorkSheet.GetCellText(counter, importSettings.ZachColumn+2) == "";

					Discipline discipline = new Discipline(counter - importSettings.StartReadingRow);
					int disciplineType = 1;
					
					discipline.Descr = disciplineName;

					//disciplineName = disciplineName.ToLower();



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
					}
					else
					{
						discipline.Special = true;

						if (disciplineName.Contains("практ"))
						{
							if (disciplineName.Contains("уч"))
							{
								discipline.UchPr = practices;
								disciplineType = 3;
							}
							else if (disciplineName.Contains("преддип"))
							{
								discipline.PredDipPr = practices;
								disciplineType = 3;
							}
							else if (disciplineName.Contains("произв"))
							{
								discipline.PrPr = practices;
								disciplineType = 3;
							}
							else if ((disciplineName.Contains("науч") && disciplineName.Contains("иссл")) 
								|| disciplineName.Contains("нир"))
							{
								discipline.NIIR = practices;
								disciplineType = 2;
							}
						}

						else if (disciplineName.Contains("конс") && disciplineName.Contains("заоч"))
						{
							discipline.KonsZaoch = true;
							disciplineType = 3;
						}

                        else if (disciplineName.Contains("гэк"))
                        {
                            discipline.GEK = true;
                            disciplineType = 3;
                        }

                        else if (disciplineName.Contains("гос") && disciplineName.Contains("экз"))
						{
							disciplineType = 3;
							discipline.GosEkz = true;
							//Добавить в таблицу и классы Гос Экзамен
						}

						else if (disciplineName.Contains("гак"))
						{
							if (disciplineName.Contains("предс"))
							{
								discipline.GAKPred = true;
							}
							else
							{
								discipline.GAK = true;
							}
							disciplineType = 3;
						}

						else if (disciplineName.Contains("вып") && disciplineName.Contains("раб"))
						{
							discipline.DPRuk = true;
							disciplineType = 2;
						}
						else if (disciplineName.Contains("диссер") )
						{
							if (disciplineName.Contains("маг"))
							{
								discipline.MAGRuk = true;
							}
							else
							{
								discipline.DPRuk = true;
							}
							disciplineType = 2;
						}

						else if (disciplineName.Contains("рук"))
						{
							if (disciplineName.Contains("каф"))
							{
								discipline.RukKaf = true;
								disciplineType = 3;
							}
							if (disciplineName.Contains("асп"))
							{
								discipline.ASPRuk = true;
								disciplineType = 2;
							}
						}
					}
					discipline.DepartmentID = 1;

					var workPlanDiscipline = new Entities.CurriculumDiscipline(_name: disciplineName,
							_semester: semester,
							_lectureCount: lectures,
							_labCount: labs,
							_practiceCount: practices,
							_courseWork: kr || kp,
							_ekz: ekz,
							_zach: zach,
							_isPractice: discipline.UchPr != 0 || discipline.PrPr != 0 || discipline.PredDipPr != 0);


					discipline.TypeID = DataManager.SharedDataManager().FindTypeByName(disciplineName);

					//year - надо чтоб: int(year) - semester/2
					int course = 1;
                    //Ввести проверку, что если семестр болльше 8(для рассчета курса магистров)
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

                    if (!Entities.Сurriculum.checkWorkPlan(workPlanDiscipline))
                    {
						answer.Add(String.Format("Дисциплина '{0}' не соответствует учебному плану\n", disciplineName));
						res = false;
                    }

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
					
					//Актуализация длительности семестра в БД
                    if(DataManager.SharedDataManager().GetSemester(semesterID).WeekCount!=weeks)
                    {
                        var newSemester = new Semester(semesterID);
                        newSemester.WeekCount = weeks;
                        newSemester.Name = semesterID.ToString();
                        DataManager.SharedDataManager().SetSemester(newSemester);
                    }
					//Актуализация информации о группе в БД
                    if (groupID != -1)
                    {
						var groupData=DataManager.SharedDataManager().GetGroup(groupID);
						bool groupChange = false;
						if(groupData.SubgroupCount != groupCount)
                        {
							groupData.SubgroupCount = groupCount;
							groupChange = true;

						}
						if(groupData.StudentCount != studentCount)
                        {
							groupData.StudentCount = studentCount;
							groupChange=true;
                        }
                        if (groupChange)
                        {
							DataManager.SharedDataManager().SetGroup(groupData);
						}
					}
					//Добавить проверку рабочего плана
					//если в дисциплине ошибка запоминать название дисциплины и выводить в МО и прерывать импорт.
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
