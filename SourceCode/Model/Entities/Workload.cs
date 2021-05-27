using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    namespace Entity
    {
        /// <summary>
        /// Нагрузка
        /// </summary>
        public class Workload
        {
            public Workload()
            {
                ID = -1;
            }

            public Workload(int id)
            {
                ID = id;
            }

            public int Discipline { get; set; }

            public int Group { get; set; }

            public int Semester { get; set; }

            public int Year { get; set; }

            public int ID { get; set; }
        }
    }
}
