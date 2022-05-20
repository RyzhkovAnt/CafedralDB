using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using TemplateEngine.Docx;
using CafedralDB.SourceCode.Model;
using CafedralDB.SourceCode.Model.Entity;

namespace CafedralDB.SourceCode.Model.Exporter
{
    /// <summary>
    /// Экспорт индивидуального плана
    ///  Необходимо проверить правильность заполнения
    /// </summary>
    public static class ExportIndPlan
    {
        public static void Export(int teacherID, string year)
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

            PrintSemester(teacherID, Utilities.SemesterName.Осенний, year, ObjWorkBook);
            PrintSemester(teacherID, Utilities.SemesterName.Весенний, year, ObjWorkBook);

            ObjExcel.Visible = true;
            ObjExcel.UserControl = true;

        }

        private static void PrintSemester(int teacherID, Utilities.SemesterName semester, string year, Excel.Workbook objWorkBook)
        {
            var semesterParam=Utilities.getSemesterConditionString(semester);
            int DisciplineDescrList;
            int DisciplineParamsList;

            if (semester == Utilities.SemesterName.Осенний)
            {
                DisciplineDescrList = 2;
                DisciplineParamsList = 4;
            }
            else
            {
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

            /*
            0 - Discipline.Descr,
            1 - Faculty.Descr AS Expr1,
            2 - Speciality.Descr AS Expr2,
            3 - [Group].EntryYear,"
            4 - [Group].StudentCount,
            5 - Semester.WeekCount,
            6 - DisciplineType.Descr AS Expr3,
            7 - StudyYear.StudyYear,
            8 - WorkloadAssign.Teacher,
            9 - Discipline.KonsZaoch,
            10 - Discipline.GAKPred,
            11 - Discipline.DPruk,
            12 - Discipline.DopuskVKR,
            13 - Discipline.RetzVKR,
            14 - Discipline.DPretz,"
            15 - Discipline.ASPRuk,
            16 - Discipline.MAGRuk,
            17 - Discipline.MAGRetz,
            18 - Discipline.RukKaf,
            19 - Discipline.NIIR,
            20 - Discipline.isSpecial,"
            21 - Workload.ID 
             
             */


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
                    var workload = DataManager.SharedDataManager().GetWorkload(Convert.ToInt32(reader[21]));

                    var discipline = DataManager.SharedDataManager().GetDiscipline(workload.Discipline);

                    string descr = string.Format("{0},{1},{2} курс", reader[1], reader[2], cource);
                    DescrSheet.Cells[currentRow, ApplicationSettings.ExportSettings.IndPlanSetting.DisciplineDescrColumn] = discipline.Descr;// descr;
                    DescrSheet.Cells[currentRow, ApplicationSettings.ExportSettings.IndPlanSetting.StudentsCountColumn] = reader[4];//studentCount
                    int countStud = Convert.ToInt32(reader[4]);
                    WorkloadCost workloadCost = Calculator.GetWorkloadCost(workload.ID);
                    Settings.CalculationSetting calculationSetting = new Settings.CalculationSetting();
                    if (discipline.Special)
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

                    var DPruk = workloadCost.DPRukCost;
                    var AspRuk = workloadCost.ASPRukCost;
                    var _rukMag = workloadCost.MagRuk;
                    
                    double rukMag = Convert.ToBoolean(reader[16]) ? (30 * Convert.ToInt32(reader[16])) : 0;//рук маг _rukMag
                    geks = GEK + GAK + GAKpred;

                    ruk += Convert.ToBoolean(reader[11]) ? (countStud * calculationSetting.DPruk) : 0f;//DPruk DPruk
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
}
