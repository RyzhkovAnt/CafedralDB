using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// Ученое звание
    /// </summary>
    public class AcademicRank
    {
        public AcademicRank(int id)
		{
			ID = id;
		}

        public string Name { get; set; }

        public int ID { get; set; }
    }
}
