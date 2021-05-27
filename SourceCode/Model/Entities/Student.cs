using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// Студент
    /// </summary>
    public class Student
    {
		int _ID;

        public Student(int id)
		{
			_ID = id;
		}

        public string Name { get; set; }
    }
}
