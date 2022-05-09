using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafedralDB.SourceCode.Settings
{
    internal  class SettingFields
    {
    }
    public struct Field
    {
        public string Name;
        public string DataBaseName;
        public object DefaultValue;

        public Field(string _name,object _defautValue, string _dataBaseName="")
        {
            Name = _name;
            DataBaseName = _dataBaseName;
            DefaultValue = _defautValue;
        }
    }
    
    public static class CalculationSettingFields {
        public static Field LectureCost = new Field("LectureCost", "Лекция");
    
    }
    public static class ImportSettingFields 
    {
        public static Field StartReadingRow = new Field ("StartReadingRow",7);

        public static Field GroupColumn = new Field("GroupColumn",1);
        public static Field StudentCountColumn = new Field("StudentCountColumn",2);
        public static Field GroupCountColumn = new Field("GroupCountColumn",5);
        public static Field SemesterColumn  = new Field("SemesterColumn",6);
        public static Field WeeksColumn  = new Field("WeeksColumn",7);
        public static Field DisciplineNameColumn = new Field("DisciplineNameColumn",8);
        public static Field LecturesColumn = new Field("LecturesColumn",10);
        public static Field PracticesColumn = new Field("PracticesColumn",11);
        public static Field LabsColumn = new Field("LabsColumn",12);
        public static Field KrColumn = new Field("KrColumn",13);
        public static Field KpColumn = new Field("KpColumn",14);
        public static Field EkzColumn = new Field("EkzColumn",15);
        public static Field ZachColumn = new Field("ZachColumn",16);
        public static Field KzColumn = new Field("KzColumn",17);
        public static Field SpecialColumn = new Field("SpecialColumn",18);

    }
}
