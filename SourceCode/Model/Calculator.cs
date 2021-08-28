using System;
using System.Collections.Generic;
using System.Reflection;
using Model.Entity;
using ApplicationSettings;
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
            lec = CalculationSettings.LectureCost * discipline.LectureCount * semester.WeekCount;
            var subGroupCount =group.SubgroupCount == 1 ? Math.Floor((double)group.StudentCount / 9)<1?1:Math.Floor((double)group.StudentCount / 9) :
                Math.Floor((double)group.StudentCount / group.SubgroupCount / 9) * group.SubgroupCount;
            lab = CalculationSettings.LabCost * discipline.LabCount * semester.WeekCount * Convert.ToInt32(subGroupCount);
                //((int)Math.Floor((double)group.StudentCount/9)>0 ? (int)Math.Floor((double)group.StudentCount / 9):1);
            prac = CalculationSettings.PracticeCost * discipline.PracticeCount * semester.WeekCount * group.SubgroupCount;
            workloadCost.LectureCost = lec; 
            workloadCost.LabCost= lab;
            workloadCost.PracticeCost= prac;

            //Расчет консультаций по дисциплине
            if (group.StudyFormID == 1)
            {
                kons = CalculationSettings.KonsCost * discipline.LectureCount * semester.WeekCount;
                workloadCost.KonsCost = kons;
            }
            else
            {
                kons = CalculationSettings.KonsCost * 3 * discipline.LectureCount * semester.WeekCount;
                workloadCost.KonsCost = kons;
            }

            if (discipline.KR)
            {
                kr = CalculationSettings.KRCost * group.StudentCount;
                workloadCost.KRCost = kr;
            }

            if (discipline.KP)
            {
                kp = CalculationSettings.KPCost * group.StudentCount;
                workloadCost.KPCost = kp;
            }

            if (discipline.Ekz)
            {
                if (group.StudyFormID == 1)
                {
                    ekz = CalculationSettings.EkzCost * group.StudentCount;
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
                zach = CalculationSettings.ZachCost * group.StudentCount;
                workloadCost.ZachCost = zach;
            }

            workloadCost.UchPracCost = CalculationSettings.UchPr * 5 * discipline.UchPr * group.SubgroupCount;
            workloadCost.PrPracCost = group.StudyFormID == 1 ? CalculationSettings.PrPr * 5 * discipline.PrPr * group.SubgroupCount :
                1 * group.StudentCount;
            workloadCost.PredDipPracCost = CalculationSettings.PreddipPr * group.StudentCount * discipline.PredDipPr;
                
            workloadCost.GEKCost = discipline.GEK ? CalculationSettings.GEK * group.StudentCount * 6 : 0;//GEK
            workloadCost.GAKCost = discipline.GAK ? CalculationSettings.GAK * group.StudentCount * 6 : 0;//GAK
            
            workloadCost.GAKPredCost = discipline.GAKPred ? CalculationSettings.GAKPred * group.StudentCount : 0;//GAKPred
            workloadCost.DPRukCost = discipline.DPRuk ? (CalculationSettings.DPruk + CalculationSettings.DopuskBak + CalculationSettings.NormocontrolBak) *
                group.StudentCount : 0;//DPruk
            workloadCost.RukKafCost = discipline.RukKaf ? CalculationSettings.RukKaf : 0;
            workloadCost.NIIRCost = CalculationSettings.NIIR * discipline.NIIR * group.StudentCount;
            workloadCost.ASPRukCost = discipline.ASPRuk ? CalculationSettings.AspRuk : 0;//GEK

            workloadCost.GosEkz = discipline.GosEkz ?
                CalculationSettings.GosEkz * group.StudentCount * CalculationSettings.EkzBoard : 0;//Гос Экзамен

            workloadCost.MagRuk = discipline.MAGRuk ? 
                group.StudentCount * CalculationSettings.MAGRuk : 0;//Маг диссер
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
