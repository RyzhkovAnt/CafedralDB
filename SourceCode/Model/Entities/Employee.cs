using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// Сотрудник
    /// </summary>
   public  class Employee
    {
        public Employee(int id)
		{
			ID = id;
		}

        public string Name { get; set; }

        internal int WorkingPositionID { get; set; }

        internal int AcademicDegreeID { get; set; }

        internal int AcademicRankID { get; set; }

        public float Rate { get; set; }

        public int ContractNumber { get; set; }

        public DateTime BirthDay { get; set; }

        public int ID { get; set; }
    }
}
