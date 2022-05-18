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
        public static Field LectureCost = new Field("LectureCost",1, "Лекция");
        public static Field LabCost = new Field("LabCost", 1, "Практическая работа");
        public static Field PracticeCost = new Field("PracticeCost", 1, "Лабораторная работа");
        public static Field KonsCost = new Field("KonsCost", 0.05f, "Консультация");
        public static Field EkzCost = new Field("EkzCost", 0.33f, "Экзамен");
        public static Field KRCost = new Field("KRCost", 2, "КР");
        public static Field KPCost = new Field("KPCost", 3, "КП");
        public static Field ZachCost = new Field("ZachCost", 0.25f, "Зачет");
        public static Field RGR = new Field("RGR", 0.25f, "РГР");
        public static Field UchPr = new Field("UchPr", 4, "Учебная практика");
        public static Field PrPr = new Field("PrPr", 1, "Производственная практика");
        public static Field PreddipPr = new Field("PreddipPr", 1, "Преддипломная практика");
        public static Field KzZaoch = new Field("KzZaoch", 0.25f, "Контрольные задания заочников");
        public static Field GEK = new Field("GEK", 0.25f, "ГЭК");
        public static Field GAK = new Field("GAK", 0.5f, "ГАК");
        public static Field GAKPred = new Field("GAKPred", 0.25f, "ГАК председатель");
        public static Field DPruk = new Field("DPruk", 14, "ДП руководство");
        public static Field DopuskVKR = new Field("DopuskVKR", 0.5f, "Допуск к ВКР");
        public static Field RetzVKR = new Field("RetzVKR", 0.25f, "Рецензия ВКР");
        public static Field DPRetz = new Field("DPRetz", 0.25f, "ДП рецензии");
        public static Field MAGRuk = new Field("MAGRuk", 35, "Магистранты руководство");
        public static Field MagRetz = new Field("MagRetz", 0.25, "Магистранты рецензирование");
        public static Field RukKaf = new Field("RukKaf", 60, "Руководство кафедрой");
        public static Field NIIR = new Field("NIIR", 3, "Научная работа");
        public static Field NIIRRukMag = new Field("NIIRRukMag", 0.25f, "Руководство НИИР магистра");
        public static Field ASPpractice = new Field("ASPpractice", 0.25f, "Практика аспирантов");
        public static Field NIIRRukAsp = new Field("NIIRRukAsp", 0.25f, "Руководство НИИР аспирантов");
        public static Field DopuskDissMag = new Field("DopuskDissMag", 1, "Допуск к защите диссертации магистрантов");
        public static Field NormocontrolMag = new Field("NormocontrolMag", 1, "Нормоконтроль  диссертации магистрантов");
        public static Field DopuskBak = new Field("DopuskBak", 0.5, "Допуск к защите бакалавров");
        public static Field NormocontrolBak = new Field("NormocontrolBak", 0.5f, "Нормоконтроль бакалавров");
        public static Field GosEkz = new Field("GosEkz", 0.5f, "Государственный экзамен");
        public static Field EkzBoard = new Field("EkzBoard", 6, "Кол-во членов экз. комиссии");
        public static Field AspRuk = new Field("AspRuk", 50, "Руковдство аспирантами");// тоже что и NIIRRukAsp?



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
