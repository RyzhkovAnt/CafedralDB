using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    public class ContractDiscr:WorkloadCost
    {
        public string Discipline { get; set; }
        public string Group { get; set; }

        public ContractDiscr(WorkloadCost workload):base()
        {
            foreach (var pr in workload.GetType().GetProperties())
                pr.SetValue(this, pr.GetValue(workload));
        }
    }
}
