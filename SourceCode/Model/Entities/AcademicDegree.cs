using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// Ученая степень
    /// </summary>
    public class AcademicDegree
    {
        public AcademicDegree(int id)
		{
			ID = id;
		}

        public string Name { get; set; }

        public int ID { get; set; }

        public override string ToString()
		{
			return String.Format("ID:{0},\nDescr:{1}",ID,Name);
		}
	}
}
