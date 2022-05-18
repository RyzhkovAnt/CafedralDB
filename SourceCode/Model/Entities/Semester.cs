using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafedralDB.SourceCode.Model
{
    /// <summary>
    /// Семестр
    /// </summary>
   public class Semester
    {
        public Semester(int id)
		{
			ID = id;
		}

        public int ID { get; set; }

        public string Name { get; set; }

        public int WeekCount { get; set; }

        public override string ToString()
        {
            return string.Format("Semester - ID:{0}, Descr:{1}, Weeks:{2}", ID, Name, WeekCount);
        }
    }
}
