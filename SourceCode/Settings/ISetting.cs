using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafedralDB.SourceCode.Settings
{
    internal interface ISetting
    {
        void LoadFromRegistry();
        void SaveToRegistry(Dictionary<Field, object> newValues);
        void ToDefault();

    }
}
