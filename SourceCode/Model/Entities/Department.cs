using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafedralDB.SourceCode.Model
{
    /// <summary>
    /// Кафедра университета
    /// </summary>
    public class Department
    {
        public Department(int id)
		{
			ID = id;
		}

        public int CodeNumber { get; set; }

        public string Description { get; set; }

        public int ID { get; set; }
    }
}
