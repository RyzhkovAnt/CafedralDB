using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// Группа
    /// </summary>
    public class Group
    {
        
        public Group(int id)
        {
            ID = id;
        }
        public int ID { get; private set; }

        public string Name { get; set; }

        internal List<int> Students { get; set; }

        internal int Faculty { get; set; }

        internal int StudyQualification { get; set; }

        internal int Speciality { get; set; }

        public int StudentCount { get; set; }

        internal int EntryYear { get; set; }

        public int StudyFormID { get; set; }

        public int SubgroupCount { get; set; }
    }
}
