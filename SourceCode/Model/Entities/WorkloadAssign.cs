using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafedralDB.SourceCode.Model
{
	public class WorkloadAssign
	{
		int _ID;

        public WorkloadAssign(int id)
		{
			_ID = id;
		}

        public int EmployeeID { get; set; }

        public int WorkloadID { get; set; }

        public int StudentsCount { get; set; }

        public int Weeks { get; set; }

        public bool IsContract { get; set; }
    }
}
