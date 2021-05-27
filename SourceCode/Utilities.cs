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
        #region Excel exstension methods
        public static string GetCellText(this Microsoft.Office.Interop.Excel.Worksheet worksheet, int row, int column)
        {
            return ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, column]).Text;
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


    }
}
