using System;
using System.Collections.Generic;
using System.Reflection;
using CafedralDB.SourceCode.Model.Entity;
using CafedralDB.SourceCode.Settings;
using ApplicationSettings;
namespace CafedralDB.SourceCode.Model
{
    public static class Calculator
	{
		public static float GetWorkloadTotalCost(int workloadID)
		{
            WorkloadCost workloadCost = GetWorkloadCost(workloadID);
            float result = 0;
            PropertyInfo[] properties = workloadCost.GetType().GetProperties();
            foreach (PropertyInfo property in properties)
                result += property.Name != "WorkLoadID" ? Convert.ToSingle( property.GetValue(workloadCost, null).ToString()) : 0;
            return result;
		}

        public static WorkloadCost GetWorkloadCost(int workloadID)
        {
            WorkloadCost workloadCost = new WorkloadCost(workloadID);
            Workload workload = DataManager.SharedDataManager().GetWorkload(workloadID);
            Discipline discipline = DataManager.SharedDataManager().GetDiscipline(workload.Discipline);
            Group group = DataManager.SharedDataManager().GetGroup(workload.Group);
            Semester semester = DataManager.SharedDataManager().GetSemester(workload.Semester);
            CalculationSetting calculationSetting = new CalculationSetting();

            float lec = 0, lab = 0, prac = 0, ekz = 0, kr = 0, kp = 0, zach = 0, kons = 0;
            int subGroupCount = GetSubGroupCount(group.StudentCount);

            lec = calculationSetting.LectureCost * discipline.LectureCount * semester.WeekCount;
            lab = calculationSetting.LabCost * discipline.LabCount * semester.WeekCount * subGroupCount;
            prac = calculationSetting.PracticeCost * discipline.PracticeCount * semester.WeekCount * group.SubgroupCount;

            workloadCost.LectureCost = lec; 
            workloadCost.LabCost= lab;
            workloadCost.PracticeCost= prac;

            //Расчет консультаций по дисциплине
            if (group.StudyFormID == 1)
            {
                kons = calculationSetting.KonsCost * discipline.LectureCount * semester.WeekCount;
                workloadCost.KonsCost = kons;
            }
            else
            {
                kons = calculationSetting.KonsCost * 3 * discipline.LectureCount * semester.WeekCount;
                workloadCost.KonsCost = kons;
            }

            if (discipline.KR)
            {
                kr = calculationSetting.KRCost * group.StudentCount;
                workloadCost.KRCost = kr;
            }

            if (discipline.KP)
            {
                kp = calculationSetting.KPCost * group.StudentCount;
                workloadCost.KPCost = kp;
            }

            if (discipline.Ekz)
            {
                if (group.StudyFormID == 1)
                {
                    ekz = calculationSetting.EkzCost * group.StudentCount;
                    workloadCost.EkzCost= ekz;
                    workloadCost.KonsCost += 2;
                }
                else
                {
                    ekz = 0.4f * group.StudentCount;
                    workloadCost.EkzCost = ekz;
                    workloadCost.KonsCost += 2;
                }
            }

            if (discipline.Zach)
            {
                zach = calculationSetting.ZachCost * group.StudentCount;
                workloadCost.ZachCost = zach;
            }

            workloadCost.UchPracCost = calculationSetting.UchPr * 5 * discipline.UchPr * group.SubgroupCount;
            workloadCost.PrPracCost = Convert.ToBoolean(discipline.PrPr)? calculationSetting.PrPr * group.StudentCount:0;
            //workloadCost.PrPracCost = group.StudyFormID == 1 ? CalculationSettings.PrPr * 5 * discipline.PrPr * group.SubgroupCount :
            //    1 * group.StudentCount;
            workloadCost.PredDipPracCost = calculationSetting.PreddipPr * group.StudentCount * discipline.PredDipPr;
                
            workloadCost.GEKCost = discipline.GEK ? calculationSetting.GEK * group.StudentCount * 6 : 0;//GEK
            workloadCost.GAKCost = discipline.GAK ? calculationSetting.GAK * group.StudentCount * 6 : 0;//GAK
            
            workloadCost.GAKPredCost = discipline.GAKPred ? calculationSetting.GAKPred * group.StudentCount : 0;//GAKPred
            workloadCost.DPRukCost = discipline.DPRuk ? (calculationSetting.DPruk + calculationSetting.DopuskBak + calculationSetting.NormocontrolBak) *
                group.StudentCount : 0;//DPruk
            workloadCost.RukKafCost = discipline.RukKaf ? calculationSetting.RukKaf : 0;
            workloadCost.NIIRCost = calculationSetting.NIIR * discipline.NIIR * group.StudentCount;
            workloadCost.ASPRukCost = discipline.ASPRuk ? calculationSetting.AspRuk : 0;//GEK

            workloadCost.GosEkz = discipline.GosEkz ?
                calculationSetting.GosEkz * group.StudentCount * calculationSetting.EkzBoard : 0;//Гос Экзамен 

            workloadCost.MagRuk = discipline.MAGRuk ? 
                group.StudentCount * calculationSetting.MAGRuk : 0;//Маг диссер 32*studentCount(old) 20+1 * studentCount (new)
            return workloadCost;
        }
        static int GetSubGroupCount(int studentCount)
        {
            if (studentCount < 18)
            {
                return 1;
            }
            return studentCount / 9;
        }
        public static float GetEmployeeAllWorkloadsCost(int employeeID, int yearID)
		{
			List<Workload> workloads = DataManager.SharedDataManager().GetEmployeeWorkloads(employeeID, yearID);
			float result=0;
			foreach(Workload workload in workloads)
			{
				result += GetWorkloadTotalCost(workload.ID);
			}
			return result;
		}

        public static double getCountOfLectureHoure(int semesterLength,int lectureCount)
        {
            var setting = new CalculationSetting();
            return setting.LectureCost * lectureCount * semesterLength;
        }
	}
}
