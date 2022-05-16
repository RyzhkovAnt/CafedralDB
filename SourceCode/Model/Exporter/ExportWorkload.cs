using System;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using CafedralDB.SourceCode.Model;

namespace CafedralDB.SourceCode.Model.Exporter
{
    public static class ExportWorkload
    {
        public static void Export(string year)
        {
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

            string newStrSQL = "SELECT Workload.ID FROM StudyYear INNER JOIN Workload ON StudyYear.ID = Workload.StudyYear " +
                "WHERE StudyYear.StudyYear = @year ORDER BY Workload.ID;";

            DataManager.SharedDataManager();
            var cn = new OleDbConnection(DataManager.Connection.ConnectionString);
            var cmd = new OleDbCommand();
            cn.Open();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = newStrSQL;
            cmd.Parameters.AddWithValue("@year", year);
            OleDbDataReader reader = cmd.ExecuteReader();

            Dictionary<Excel.Worksheet, int> rowCounters = new Dictionary<Excel.Worksheet, int>();
            //пишем данные
            for (int i = 2; i <= 17; i++)
            {
                rowCounters.Add(ObjWorkBook.Sheets[i], 6);
            }
            var k = 0;
            while (reader.Read())
            {
                //int workloadID = Convert.ToInt32(reader[1]);
                k++;
                var workloadId = Convert.ToInt32(reader[0]);
                
                var workloadCost = Calculator.GetWorkloadCost(workloadId);
                var workload = DataManager.SharedDataManager().GetWorkload(workloadId);
                var group = DataManager.SharedDataManager().GetGroup(workload.Group);
                var speciality = DataManager.SharedDataManager().GetSpeciality(group.Speciality);
                var faculty = DataManager.SharedDataManager().GetFaculty(speciality.Faculty);
                var discipline = DataManager.SharedDataManager().GetDiscipline(workload.Discipline);
                var semester = DataManager.SharedDataManager().GetSemester(workload.Semester);

                List<WorkloadAssign> assigns = DataManager.SharedDataManager().GetWorkloadAssigns(workloadId);
                //Определим листы для записи
                List<Excel.Worksheet> toWriteList = new List<Excel.Worksheet>();
                if (faculty.Name == "МСФ")
                {
                    toWriteList.Add(ObjWorkBook.Sheets[6]);
                }
                else
                {
                    if (group.StudyFormID == 2)
                    {
                        toWriteList.Add(ObjWorkBook.Sheets[5]);
                    }
                    else
                    {
                        if (group.StudyQualification == 1)
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

                if (assigns.Count != 0)
                {
                    foreach (WorkloadAssign assign in assigns)
                    {
                        string teacherName = DataManager.SharedDataManager().GetEmployee(assign.EmployeeID).Name;
                        Excel.Worksheet sheet = ObjWorkBook.Sheets[teacherName];
                        
                        if (sheet != null)
                        {
                            toWriteList.Add(sheet);
                            assignsForTeachers.Add(sheet, assign);
                        }
                    }
                }
                else
                {
                    if (group.StudyQualification == 1)
                    {
                        toWriteList.Add(ObjWorkBook.Sheets[2]);
                    }
                    else
                    {
                        toWriteList.Add(ObjWorkBook.Sheets[3]);
                    }
                }

                foreach (Excel.Worksheet sheet in toWriteList)
                {
                    //int countStud = Convert.ToInt32(reader[8]);
                    //if (sheet.Index > 8 && Convert.ToBoolean(reader[34]))
                    //{
                    //    WorkloadAssign assign = DataManager.SharedDataManager().
                    //        GetWorkloadAssign(Convert.ToInt32(reader[1]), sheet.Name);
                    //    countStud = assign.StudentsCount;

                    //}
                    //TestWorkload(Convert.ToInt32(reader[1]));/*DEBUG*/
                    ExportNotAssign(workloadCost, group, speciality, faculty, discipline, semester, sheet, rowCounters);
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

        private static void ExportNotAssign(WorkloadCost workloadCost,Group group,Speciality speciality,Faculty faculty,Discipline discipline,

            CafedralDB.SourceCode.Model.Semester semester, Excel.Worksheet sheet, Dictionary<Excel.Worksheet, int> rowCounters)
        {
            //var workloadId = Convert.ToInt32(reader[1]);
            //var workloadCost = Calculator.GetWorkloadCost(Convert.ToInt32(reader[1]));
            //var workload = DataManager.SharedDataManager().GetWorkload(workloadId);
            //var group = DataManager.SharedDataManager().GetGroup(workload.Group);
            //var speciality = DataManager.SharedDataManager().GetSpeciality(group.Speciality);
            //var faculty = DataManager.SharedDataManager().GetFaculty(group.Faculty);
            //var discipline = DataManager.SharedDataManager().GetDiscipline(workload.Discipline);
            //var semester = DataManager.SharedDataManager().GetSemester(workload.Semester);

            CafedralDB.SourceCode.Settings.CalculationSetting calculationSetting = new CafedralDB.SourceCode.Settings.CalculationSetting();
            int groupCount = group.SubgroupCount;
            int studentCount = group.StudentCount;
            var semesterNumber = Convert.ToInt32(semester.Name);
            var subGroupCount = groupCount == 1 ? Math.Floor((double)studentCount / 9) > 0 ? Math.Floor((double)studentCount / 9) : 1 :
                Math.Floor((double)studentCount / groupCount / 9) * groupCount;

            sheet.Cells[rowCounters[sheet], 1] = rowCounters[sheet] - 5;//
            sheet.Cells[rowCounters[sheet], 2] = discipline.Descr;//Дисциплина
            sheet.Cells[rowCounters[sheet], 3] = speciality.Name;//Специальность
            sheet.Cells[rowCounters[sheet], 4] = faculty.Name;//Факультет
            sheet.Cells[rowCounters[sheet], 5] = group.StudyFormID == 1 ? "о" : "з";//Форма обучения
            sheet.Cells[rowCounters[sheet], 6] = semesterNumber > 8 ? semesterNumber - 8 : semesterNumber;//Семестр
            sheet.Cells[rowCounters[sheet], 7] = group.StudentCount ;//reader[8].ToString(); Кол-во студентов
            sheet.Cells[rowCounters[sheet], 8] = semester.WeekCount; // Кол-во недель
            sheet.Cells[rowCounters[sheet], 9] = group.SubgroupCount; //reader[10].ToString(); // Кол-во групп
            sheet.Cells[rowCounters[sheet], 10] = subGroupCount; // Кол-во подгрупп
            sheet.Cells[rowCounters[sheet], 11] = discipline.LectureCount; // Кол-во лекций
            sheet.Cells[rowCounters[sheet], 12] = discipline.PracticeCount; // Кол-во практ. работ
            sheet.Cells[rowCounters[sheet], 13] = discipline.LabCount; // Кол-во лаб. работ
            sheet.Cells[rowCounters[sheet], 14] = discipline.Ekz ? "+" : "";//Экзамен
            sheet.Cells[rowCounters[sheet], 15] = discipline.Zach ? "+" : "";// Зачет
            sheet.Cells[rowCounters[sheet], 16] = discipline.RGR ? "+" : "";// РГр
            sheet.Cells[rowCounters[sheet], 17] = discipline.KR ? "+" : "";// КР
            sheet.Cells[rowCounters[sheet], 18] = discipline.KP ? "+" : "";// КП
            sheet.Cells[rowCounters[sheet], 19] = Convert.ToBoolean(discipline.UchPr) ? "+" : "";// Учебная практика
            sheet.Cells[rowCounters[sheet], 20] = Convert.ToBoolean(discipline.PrPr) ? "+" : "";// Производственная практика
            sheet.Cells[rowCounters[sheet], 21] = Convert.ToBoolean(discipline.PredDipPr) ? "+" : "";// Преддипломная практика
            sheet.Cells[rowCounters[sheet], 23] = discipline.GEK ? "+" : "";// ГЭК
            sheet.Cells[rowCounters[sheet], 24] = discipline.GAK ? "+" : "";// Гак
            sheet.Cells[rowCounters[sheet], 25] = discipline.GAKPred ? "+" : "";// Гак председатель
            sheet.Cells[rowCounters[sheet], 33] = discipline.RukKaf ? "+" : "";// Руководство кафедрой


            /*
                Заполнение часов начинается с ячейки 34 и имеет след порядок:
                    - лекции - 34
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
            sheet.Cells[rowCounters[sheet], 34] = Math.Round(workloadCost.LectureCost, 2); //часов лекций
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
                sheet.Cells[rowCounters[sheet], 8] = discipline.UchPr; // Кол-во недель
                sheet.Cells[rowCounters[sheet], 19] = Convert.ToBoolean(discipline.UchPr) ? "+" : "";
                sheet.Cells[rowCounters[sheet], 43] = Math.Round(workloadCost.UchPracCost, 2);
            }
            //Производственная практика
            if (workloadCost.PrPracCost != 0)
            {
                sheet.Cells[rowCounters[sheet], 8] = discipline.PrPr; // Кол-во недель
                sheet.Cells[rowCounters[sheet], 20] = Convert.ToBoolean(discipline.PrPr) ? "+" : "";
                sheet.Cells[rowCounters[sheet], 44] = Math.Round(workloadCost.PrPracCost, 2);
            }
            //Преддипломная практика
            if (workloadCost.PredDipPracCost != 0)
            {
                sheet.Cells[rowCounters[sheet], 8] = discipline.PredDipPr; // Кол-во недель
                sheet.Cells[rowCounters[sheet], 21] = Convert.ToBoolean(discipline.PredDipPr) ? "+" : "";
                sheet.Cells[rowCounters[sheet], 45] = Math.Round(workloadCost.PredDipPracCost, 2);
            }
            //Научно-исследовательская работа
            if (workloadCost.NIIRCost != 0)
            {
                sheet.Cells[rowCounters[sheet], 8] = discipline.NIIR; // Кол-во недель
                sheet.Cells[rowCounters[sheet], 22] = Convert.ToBoolean(discipline.NIIR) ? "+" : "";
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
                sheet.Cells[rowCounters[sheet], 50] = Math.Round(workloadCost.DPRukCost, 2);
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

            if (discipline.DopuskVKR)
            {
                sheet.Cells[rowCounters[sheet], 51] = (4 * group.StudentCount).ToString();
            }
            //Руководство аспирантами
            if (discipline.ASPRuk)
            {
                sheet.Cells[rowCounters[sheet], 54] = (group.StudentCount * calculationSetting.AspRuk);
            }

            //sheet.Cells[rowCounters[sheet], 51] = Convert.ToBoolean(reader[26]) ?
            //    (4 * group.StudentCount).ToString() : "";//рец дисс.

            //sheet.Cells[rowCounters[sheet], 54] = Convert.ToBoolean(reader[29]) ?
            //    (group.StudentCount * ApplicationSettings.CalculationSettings.AspRuk) : 0f; //AspRuk

            rowCounters[sheet]++;
        }
    }
}

