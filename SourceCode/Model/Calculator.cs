using System;
using System.Collections.Generic;
using System.Reflection;
using Model.Entity;

namespace Model
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
            float lec = 0, lab = 0, prac = 0, ekz = 0, kr = 0, kp = 0, zach = 0, kons = 0;
            lec = ApplicationSettings.CalculationSettings.LectureCost * discipline.LectureCount * semester.WeekCount;
            lab = ApplicationSettings.CalculationSettings.LabCost * discipline.LabCount * semester.WeekCount * 
                ((int)Math.Floor((double)group.StudentCount/9)>0 ? (int)Math.Floor((double)group.StudentCount / 9):1);
            prac = ApplicationSettings.CalculationSettings.PracticeCost * discipline.PracticeCount * semester.WeekCount * group.SubgroupCount;
            workloadCost.LectureCost = lec; 
            workloadCost.LabCost= lab;
            workloadCost.PracticeCost= prac;

            //Расчет консультаций по дисциплине
            if (group.StudyFormID == 1)
            {
                kons = ApplicationSettings.CalculationSettings.KonsCost * discipline.LectureCount * semester.WeekCount;
                workloadCost.KonsCost = kons;
            }
            else
            {
                kons = ApplicationSettings.CalculationSettings.KonsCost * 3 * discipline.LectureCount * semester.WeekCount;
                workloadCost.KonsCost = kons;
            }

            if (discipline.KR)
            {
                kr = ApplicationSettings.CalculationSettings.KRCost * group.StudentCount;
                workloadCost.KRCost = kr;
            }

            if (discipline.KP)
            {
                kp = ApplicationSettings.CalculationSettings.KPCost * group.StudentCount;
                workloadCost.KPCost = kp;
            }

            if (discipline.Ekz)
            {
                if (group.StudyFormID == 1)
                {
                    ekz = ApplicationSettings.CalculationSettings.EkzCost * group.StudentCount;
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
                zach = ApplicationSettings.CalculationSettings.ZachCost * group.StudentCount;
                workloadCost.ZachCost = zach;
            }
           
            workloadCost.UchPracCost = ApplicationSettings.CalculationSettings.UchPr * discipline.UchPr;
            workloadCost.PrPracCost = ApplicationSettings.CalculationSettings.PrPr * discipline.PrPr;
            workloadCost.PredDipPracCost = group.StudyFormID == 1 ? ApplicationSettings.CalculationSettings.PreddipPr * 5 * discipline.PredDipPr * group.SubgroupCount :
                ApplicationSettings.CalculationSettings.PreddipPr * group.StudentCount;
                
            //ApplicationSettings.CalculationSettings.PreddipPr * discipline.PredDipPr;
            workloadCost.GEKCost = discipline.GEK ? ApplicationSettings.CalculationSettings.GEK * group.StudentCount * 6 : 0;//GEK
            workloadCost.GAKCost = discipline.GAK ? ApplicationSettings.CalculationSettings.GAK * group.StudentCount * 6 : 0;//GAK
            
            workloadCost.GAKPredCost = discipline.GAKPred ? ApplicationSettings.CalculationSettings.GAKPred * group.StudentCount : 0;//GAKPred
            workloadCost.DPRukCost = discipline.DPRuk ? ApplicationSettings.CalculationSettings.DPruk * group.StudentCount : 0;//DPruk
            workloadCost.RukKafCost = discipline.RukKaf ? ApplicationSettings.CalculationSettings.RukKaf : 0;
            workloadCost.NIIRCost = ApplicationSettings.CalculationSettings.NIIR * discipline.NIIR * group.StudentCount;
            workloadCost.ASPRukCost = discipline.ASPRuk ? ApplicationSettings.CalculationSettings.AspRuk : 0;//GEK

            return workloadCost;
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
	}
}
