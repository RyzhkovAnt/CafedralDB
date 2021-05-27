using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// Учебный год
    /// </summary>
    public class StudyYear
    {
        public StudyYear(int id)
		{
			ID = id;
		}

        public int ID { get; set; }

        public int Year { get; set; }

        public override string ToString()
        {
            return string.Format("{0}-{1}",Year,Year+1);
        }
    }
}
