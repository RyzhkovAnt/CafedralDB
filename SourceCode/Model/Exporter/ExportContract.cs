using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using TemplateEngine.Docx;

namespace CafedralDB.SourceCode.Model.Exporter
{
    public static class ExportContract
    {
        public static void Export(int teacherID, string year, Utilities.SemesterName semester)
        {

            string strSQL = "SELECT Workload.ID, Group.Descr, Discipline.Descr " +
                "FROM(StudyYear INNER JOIN(Semester INNER JOIN([Group] " +
                "INNER JOIN(Discipline INNER JOIN Workload ON Discipline.ID = Workload.Discipline) ON " +
                "Group.ID = Workload.Group) ON Semester.ID = Workload.Semester) ON " +
                "StudyYear.ID = Workload.StudyYear) INNER JOIN(Employee INNER JOIN " +
                "WorkloadAssign ON Employee.ID = WorkloadAssign.Teacher) ON Workload.ID = WorkloadAssign.Workload " +
                "WHERE ((WorkloadAssign.isContract) = True) AND((StudyYear.StudyYear) =[param1]) AND((WorkloadAssign.Teacher) =[@param2]) AND " + Utilities.getSemesterConditionString(semester);
            DataManager.SharedDataManager();
            var cn = new OleDbConnection(DataManager.Connection.ConnectionString);
            var cmd = new OleDbCommand();
            cn.Open();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strSQL;
            cmd.Parameters.AddWithValue("@param1", year);
            cmd.Parameters.AddWithValue("@param2", teacherID);
            OleDbDataReader reader = cmd.ExecuteReader();
            string teacherName = DataManager.SharedDataManager().GetEmployee(teacherID).Name;

            List<ContractDiscr> DisciplineList = new List<ContractDiscr>();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    WorkloadCost workloadCost = Calculator.GetWorkloadCost(Convert.ToInt32(reader[0]));
                    ContractDiscr contractDescr = new ContractDiscr(workloadCost);
                    contractDescr.Group = reader[1].ToString();
                    contractDescr.Discipline = reader[2].ToString();
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
            string dateWithTime = string.Format("{0}-{1}", nowDate, nowTime.ToString(@"hh\.mm\.ss"));
            string filename = string.Format("Contract of {0} - {1}({2}).docx", TeacherName, year, dateWithTime);
            string outpath = "Contracts\\" + filename;
            string[] prefixTags = ApplicationSettings.ExportSettings.ContractSetting.GetPrefixTag();
            File.Copy(path, outpath);
            TableContent tableContent = new TableContent("Discipline table");
            for (int i = 0; i < 7; i++)
            {
                tableContent.AddRow(
                    new FieldContent("row number", (i + 1).ToString()),
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

            var _path = System.Windows.Forms.Application.StartupPath + "\\" + outpath;
            System.Diagnostics.Process.Start(_path);
        }

        static List<FieldContent> FillTableOfTrainingTask(List<ContractDiscr> contractDiscrs, string[] prefixTags)
        {
            List<FieldContent> fieldContents = new List<FieldContent>();
            double totalAll = 0, totalColumn = 0;
            double value;
            string tempValue = "";
            double[] totalRow = new double[prefixTags.Length];
            string tag = "";
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < prefixTags.Length; j++)
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
            fieldContents.Add(new FieldContent(tag, totalAll.ToString()));

            return fieldContents;
        }

        static string FindItem(List<ContractDiscr> list, int index, string field)
        {
            switch (field)
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

