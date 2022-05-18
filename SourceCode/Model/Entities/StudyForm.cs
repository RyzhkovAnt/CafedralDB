using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafedralDB.SourceCode.Model
{
    /// <summary>
    /// Форма обучения
    /// </summary>
    enum StudyFormEnum
    {
        FullTime, //Очное обучение
        CorrespondenceTime, //Заочное обучение
        PartTime //Очно-заочное обучение
    }

	/// <summary>
	/// Форма обучения
	/// </summary>
	public class StudyForm
	{
        public StudyForm(int id)
		{
			ID = id;
		}

        public int ID { get; set; }

        public string Name { get; set; }

        public string NameRus { get; set; }
    }
}
