using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafedralDB.SourceCode.Model
{
    /// <summary>
    /// Рабочая должность
    /// </summary>
    public class WorkingPosition
    {
        public WorkingPosition(int id)
		{
			ID = id;
		}

        public string Name { get; set; }

        public int ID { get; set; }
    }
}
