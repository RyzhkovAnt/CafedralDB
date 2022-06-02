using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafedralDB.SourceCode
{
    public static class Utilities
    {
        public enum SemesterName { Осенний,Весенний }
        #region Excel exstension methods
        public static string GetCellText(this Microsoft.Office.Interop.Excel.Worksheet worksheet, int row, int column)
        {
            return  Convert.ToString(((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, column]).Value);
        }

        public static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        #endregion

        public static string getSemesterConditionString(SemesterName semesterName)
        {
            if (semesterName == SemesterName.Осенний)
            {
                return "(Semester.Descr = '1'  OR  Semester.Descr = '3'  OR  Semester.Descr = '5'  OR  Semester.Descr = '7'  OR  Semester.Descr = '9'  OR  Semester.Descr = '11')";
            }

            return "(Semester.Descr = '2'  OR  Semester.Descr = '4'  OR  Semester.Descr = '6'  OR Semester.Descr = '8'  OR  Semester.Descr = '10'  OR  Semester.Descr = '12')";

        }

        
    }

    
}
