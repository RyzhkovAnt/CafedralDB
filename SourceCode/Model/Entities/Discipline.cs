using System.Collections.Generic;

namespace Model
{
    /// <summary>
    /// Дисциплина
    /// </summary>
    public class Discipline
    {
        public Discipline(int id)
		{
			ID = id;
		}

        public string Descr { get; set; }

        public List<DisciplineParameter> Parameters { get; set; }
        internal DisciplineTypeEnum Type { get; set; }
        internal int DepartmentID { get; set; }
        public int TypeID { get; set; }
        public int LectureCount { get; set; }
        public int PracticeCount { get; set; }
        public int LabCount { get; set; }
        public bool KR { get; set; }
        public bool KP { get; set; }
        public bool RGR { get; set; }
        public bool Zach { get; set; }
        public bool Ekz { get; set; }
        public bool Kons { get; set; }
        public int ID { get; set; }
        public bool Kz { get; set; }
        public int SemesterID { get; set; }
        public int SpecialityID { get; set; }
        public int UchPr { get; set; }
        public int PrPr { get; set; }
        public int PredDipPr { get; set; }
        public bool KonsZaoch { get; set; }
        public bool GEK { get; set; }
        public bool GAK { get; set; }
        public bool GAKPred { get; set; }
        public bool DPRuk { get; set; }
        public bool DopuskVKR { get; set; }
        public bool RetzVKR { get; set; }
        public bool DPretz { get; set; }
        public bool ASPRuk { get; set; }
        public bool MAGRuk { get; set; }
        public bool MAGRetz { get; set; }
        public bool RukKaf { get; set; }
        public int NIIR { get; set; }
        public bool Special { get; set; } = false;
        public bool Contract { get; set; }

        public bool GosEkz { get; set; }
        public override string ToString()
		{
			string res = string.Format("ID:{0}\nDescr:{1}\nLect:{2}\nPract:{3}\nLab:{4}",
			ID,Descr,LectureCount,PracticeCount,LabCount);
			return res;
		}
	}
}
