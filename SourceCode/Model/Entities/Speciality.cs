using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// Специальность
    /// </summary>
    public class Speciality
    {
		int _ID;

        public Speciality(int id)
		{
			_ID = id;
		}

        public string Name { get; set; }

        public int Faculty { get; set; }
    }
}
