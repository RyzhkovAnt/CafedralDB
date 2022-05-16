using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafedralDB.SourceCode.Model
{
    public class Faculty
    {
		int _ID;

        public Faculty(int id)
		{
			_ID = id;
		}

        public string Name { get; set; }

        public string FullName { get; set; }
    }
}
