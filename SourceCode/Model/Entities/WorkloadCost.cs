using System;

namespace Model
{
    public class WorkloadCost
    {
        public WorkloadCost(int id)
        {
            WorkLoadID = id;
        }
        public WorkloadCost()
        {
            WorkLoadID = -1;
        }

        public int WorkLoadID{get;set;}
        public float LectureCost { get; set; }
        public float LabCost { get; set; }
        public float PracticeCost { get; set; }
        public float EkzCost { get; set; }
        public float KRCost { get; set; }
        public float KPCost { get; set; }
        public float ZachCost { get; set; }
        public float KonsCost { get; set; }
        public float UchPracCost { get; set; }
        public float PrPracCost { get; set; }
        public float PredDipPracCost { get; set; }
        public float ASPRukCost { get; set; }
        public float GEKCost { get; set; }
        public float GAKCost { get; set; }
        public float GAKPredCost { get; set; }
        public float DPRukCost { get; set; }
        public float RukKafCost { get; set; }
        public float NIIRCost { get; set; }
    }
}
