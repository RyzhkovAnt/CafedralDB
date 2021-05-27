using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// Квалификация
    /// </summary>
    public enum StudyQualification
    {
        Bachelor, //Бакалавриат
        Specialist, //Специалитет
        Magistracy //Магистратура
    }

	public class Qualification
	{
		int _ID;

        public Qualification(int id)
		{
			_ID = id;
		}

        public string Name { get; set; }
    }

}
