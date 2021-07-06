using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Windows.Forms;


namespace ApplicationSettings
{
    public static class ImportSettings
    {

        public static int StartReadingRow = 7;
        public static int GroupColumn = 1;
        public static int GroupCountColumn = 5;
        public static int StudentCountColumn = 2;

        public static int SemesterColumn = 6;
        public static int WeeksColumn = 7;
        public static int DisciplineNameColumn = 8;
        public static int LecturesColumn = 10;

        public static int PracticesColumn = 11;
        public static int LabsColumn = 12;
        public static int KzColumn = 17;
        public static int KrColumn = 13;
        public static int KpColumn = 14;
        public static int EkzColumn = 15;
        public static int ZachColumn = 16;
        public static int OtherColumn = 9;
        public static int ContractColumn = 19;
        public static int SpecialColumn = 18;

        public static void ToRegistry()
        {
            try
            {
                RegistryKey saveKey = Registry.CurrentUser.CreateSubKey("software\\ChairDB\\Import");
                saveKey.SetValue("StartReadingRow", StartReadingRow);
                saveKey.SetValue("GroupColumn", GroupColumn);
                saveKey.SetValue("GroupCountColumn", GroupCountColumn);
                saveKey.SetValue("WeeksColumn", WeeksColumn);
                saveKey.SetValue("DisciplineNameColumn", DisciplineNameColumn);
                saveKey.SetValue("LecturesColumn", LecturesColumn);
                saveKey.SetValue("LabsColumn", LabsColumn);
                saveKey.SetValue("SemesterColumn", SemesterColumn);
                saveKey.SetValue("PracticesColumn", PracticesColumn);
                saveKey.SetValue("KzColumn", KzColumn);
                saveKey.SetValue("KrColumn", KrColumn);
                saveKey.SetValue("KpColumn", KpColumn);
                saveKey.SetValue("EkzColumn", EkzColumn);
                saveKey.SetValue("ZachColumn", ZachColumn);
                saveKey.SetValue("OtherColumn", OtherColumn);
                saveKey.SetValue("SpecialColumn", SpecialColumn);

                saveKey.Close();
            }
            catch(Exception err)
            {
                MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void FromRegistry()
        {
            try
            {
                RegistryKey readKey = Registry.CurrentUser.OpenSubKey(@"software\ChairDB\Import");
                if (readKey != null)
                {
                    StartReadingRow = Convert.ToInt32(readKey.GetValue("StartReadingRow"));
                    GroupColumn = Convert.ToInt32(readKey.GetValue("GroupColumn"));
                    GroupCountColumn = Convert.ToInt32(readKey.GetValue("GroupCountColumn"));
                    SemesterColumn = Convert.ToInt32(readKey.GetValue("SemesterColumn"));
                    WeeksColumn = Convert.ToInt32(readKey.GetValue("WeeksColumn"));
                    DisciplineNameColumn = Convert.ToInt32(readKey.GetValue("DisciplineNameColumn"));
                    LecturesColumn = Convert.ToInt32(readKey.GetValue("LecturesColumn"));
                    LabsColumn = Convert.ToInt32(readKey.GetValue("LabsColumn"));
                    PracticesColumn = Convert.ToInt32(readKey.GetValue("PracticesColumn"));
                    KzColumn = Convert.ToInt32(readKey.GetValue("KzColumn"));
                    KrColumn = Convert.ToInt32(readKey.GetValue("KrColumn"));
                    KpColumn = Convert.ToInt32(readKey.GetValue("KpColumn"));
                    EkzColumn = Convert.ToInt32(readKey.GetValue("EkzColumn"));
                    ZachColumn = Convert.ToInt32(readKey.GetValue("ZachColumn"));
                    OtherColumn = Convert.ToInt32(readKey.GetValue("OtherColumn"));
                    SpecialColumn = Convert.ToInt32(readKey.GetValue("SpecialColumn"));

                    readKey.Close();
                }

            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void ToDefault()
        {
            try
            {
                RegistryKey saveKey = Registry.CurrentUser.CreateSubKey("software\\ChairDB\\Import");
                saveKey.SetValue("StartReadingRow", 7);
                saveKey.SetValue("GroupColumn", 1);
                saveKey.SetValue("GroupCountColumn", 5);
                saveKey.SetValue("WeeksColumn", 7);
                saveKey.SetValue("DisciplineNameColumn", 8);
                saveKey.SetValue("LecturesColumn", 10);
                saveKey.SetValue("LabsColumn", 12);
                saveKey.SetValue("SemesterColumn", 6);
                saveKey.SetValue("PracticesColumn", 11);
                saveKey.SetValue("KzColumn", 17);
                saveKey.SetValue("KrColumn", 13);
                saveKey.SetValue("KpColumn", 14);
                saveKey.SetValue("EkzColumn", 15);
                saveKey.SetValue("ZachColumn", 16);
                saveKey.SetValue("OtherColumn", 9);
                saveKey.SetValue("ContractColumn", 19);
                saveKey.SetValue("SpecialColumn", 18);

                saveKey.Close();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    public static class CalculationSettings
    {
        public static float LectureCost = 1;
        public static float LabCost = 1;
        public static float PracticeCost = 1;
        public static float KonsCost = 0.05f;
        public static float EkzCost = 0.33f;
        public static float KRCost = 2;
        public static float KPCost = 3;
        public static float ZachCost = 0.25f;

        public static float RGR = 0.25f;
        public static float UchPr = 4;
        public static float PrPr = 2;
        public static float PreddipPr = 1;
        public static float KzZaoch = 0.25f;
        public static float GEK = 0.25f;
        public static float GAK = 0.5f;
        public static float GAKPred = 0.25f;
        public static float DPruk = 14;
        public static float DopuskVKR = 0.5f;
        public static float RetzVKR = 0.25f;
        public static float DPRetz = 0.25f;
        public static float MAGRuk = 36;
        public static float MagRetz = 0.25f;
        public static float RukKaf = 60;
        public static float NIIR = 3.0f;
        public static float NIIRRukMag = 0.25f;
        public static float ASPpractice = 0.25f;
        public static float NIIRRukAsp = 0.25f;
        public static float DopuskDissMag = 1;
        public static float NormocontrolMag = 1;
        public static float DopuskBak = 0.5f;
        public static float NormocontrolBak = 0.5f;
        public static float AspRuk = 50f;
        public static float GosEkz = 0.5f;
        public static float EkzBoard = 6;
        public static void ToRegistry()
        {
            try
            {
                RegistryKey saveKey = Registry.CurrentUser.CreateSubKey(@"software\ChairDB\CalculationSettings");
                saveKey.SetValue("LectureCost", LectureCost);
                saveKey.SetValue("LabCost", LabCost);
                saveKey.SetValue("PracticeCost", PracticeCost);
                saveKey.SetValue("KonsCost", KonsCost);
                saveKey.SetValue("EkzCost", EkzCost);
                saveKey.SetValue("KRCost", KRCost);
                saveKey.SetValue("KPCost", KPCost);
                saveKey.SetValue("ZachCost", ZachCost);

                saveKey.SetValue("RGR", RGR);
                saveKey.SetValue("UchPr", UchPr);
                saveKey.SetValue("PrPr", PrPr);
                saveKey.SetValue("PreddipPr", PreddipPr);
                saveKey.SetValue("KzZaoch", KzZaoch);
                saveKey.SetValue("GEK", GEK);
                saveKey.SetValue("GAK", GAK);
                saveKey.SetValue("GAKPred", GAKPred);
                saveKey.SetValue("DPruk", DPruk);
                saveKey.SetValue("DopuskVKR", DopuskVKR);
                saveKey.SetValue("RetzVKR", RetzVKR);
                saveKey.SetValue("DPRetz", DPRetz);
                saveKey.SetValue("MAGRuk", MAGRuk);
                saveKey.SetValue("MagRetz", MagRetz);
                saveKey.SetValue("RukKaf", RukKaf);
                saveKey.SetValue("NIIR", NIIR);
                saveKey.SetValue("NIIRRukMag", NIIRRukMag);
                saveKey.SetValue("ASPpractice", ASPpractice);
                saveKey.SetValue("NIIRRukAsp", NIIRRukAsp);
                saveKey.SetValue("DopuskDissMag", DopuskDissMag);
                saveKey.SetValue("NormocontrolMag", NormocontrolMag);
                saveKey.SetValue("DopuskBak", DopuskBak);
                saveKey.SetValue("NormocontrolBak", NormocontrolBak);
                saveKey.SetValue("AspRuk", AspRuk);
                saveKey.SetValue("GosEkz", GosEkz);
                saveKey.SetValue("EkzBoard", EkzBoard);

                saveKey.Close();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void FromRegistry()
        {
            try
            {
                RegistryKey readKey = Registry.CurrentUser.OpenSubKey(@"software\ChairDB\CalculationSettings");
                if (readKey != null)
                {
                    LectureCost = Convert.ToSingle(readKey.GetValue("LectureCost"));
                    LabCost = Convert.ToSingle(readKey.GetValue("LabCost"));
                    PracticeCost = Convert.ToSingle(readKey.GetValue("PracticeCost"));
                    KonsCost = Convert.ToSingle(readKey.GetValue("KonsCost"));
                    EkzCost = Convert.ToSingle(readKey.GetValue("EkzCost"));
                    KRCost = Convert.ToSingle(readKey.GetValue("KRCost"));
                    KPCost = Convert.ToSingle(readKey.GetValue("KPCost"));
                    ZachCost = Convert.ToSingle(readKey.GetValue("ZachCost"));

                    RGR = Convert.ToSingle(readKey.GetValue("RGR"));
                    UchPr = Convert.ToInt32(readKey.GetValue("UchPr"));
                    PrPr = Convert.ToInt32(readKey.GetValue("PrPr"));
                    PreddipPr = Convert.ToSingle(readKey.GetValue("PreddipPr"));
                    KzZaoch = Convert.ToSingle(readKey.GetValue("KzZaoch"));
                    GEK = Convert.ToSingle(readKey.GetValue("GEK"));
                    GAK = Convert.ToSingle(readKey.GetValue("GAK"));
                    GAKPred = Convert.ToSingle(readKey.GetValue("GAKPred"));
                    DPruk = Convert.ToSingle(readKey.GetValue("DPruk"));
                    DopuskVKR = Convert.ToSingle(readKey.GetValue("DopuskVKR"));
                    RetzVKR = Convert.ToSingle(readKey.GetValue("RetzVKR"));
                    DPRetz = Convert.ToSingle(readKey.GetValue("DPRetz"));
                    MAGRuk = Convert.ToSingle(readKey.GetValue("MAGRuk"));
                    MagRetz = Convert.ToSingle(readKey.GetValue("MagRetz"));
                    RukKaf = Convert.ToSingle(readKey.GetValue("RukKaf"));
                    NIIR = Convert.ToSingle(readKey.GetValue("NIIR"));
                    NIIRRukMag = Convert.ToSingle(readKey.GetValue("NIIRRukMag"));
                    ASPpractice = Convert.ToSingle(readKey.GetValue("ASPpractice"));
                    NIIRRukAsp = Convert.ToSingle(readKey.GetValue("NIIRRukAsp"));
                    DopuskDissMag = Convert.ToSingle(readKey.GetValue("DopuskDissMag"));
                    NormocontrolMag = Convert.ToSingle(readKey.GetValue("NormocontrolMag"));
                    DopuskBak = Convert.ToSingle(readKey.GetValue("DopuskBak"));
                    NormocontrolBak = Convert.ToSingle(readKey.GetValue("NormocontrolBak"));
                    GosEkz = Convert.ToSingle(readKey.GetValue("GosEkz"));
                    EkzBoard = Convert.ToInt32(readKey.GetValue("EkzBoard"));
                    readKey.Close();
                }
                else
                    FromDataBase();
                
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "CalculationSettings.FromRegistry", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void FromDataBase()
        {
            try
            {
                bool res = false;
                Model.DataManager.SharedDataManager();
                var cn = new System.Data.OleDb.OleDbConnection(Model.DataManager.Connection.ConnectionString);
                var cmd = new System.Data.OleDb.OleDbCommand();
                cn.Open();
                cmd.Connection = cn;
                cmd.CommandType = System.Data.CommandType.Text;
                cmd.CommandText = "Select ParameterName, Value From Normas";
                System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
                // while there is another record present
                res = reader.HasRows;
                Dictionary<string, float> parameters = new Dictionary<string, float>();
                if (res)
                {
                    while (reader.Read())
                    {
                        string name = reader[0].ToString();
                        float value = Convert.ToSingle(reader[1]);
                        parameters.Add(name, value);
                    }
                }

                LectureCost = parameters["Лекция"];
                LabCost = parameters["Практическая работа"];
                PracticeCost = parameters["Лабораторная работа"];
                KonsCost = parameters["Консультация"];
                EkzCost = parameters["Экзамен"];
                KRCost = parameters["КР"];
                KPCost = parameters["КП"];
                ZachCost = parameters["Зачет"];
                RGR = parameters["РГР"];
                UchPr = parameters["Учебная практика"];
                PrPr = parameters["Производственная практика"];
                PreddipPr = parameters["Преддипломная практика"];
                KzZaoch = parameters["Контрольные задания заочников"];
                GEK = parameters["ГЭК"];
                GAK = parameters["ГАК"];
                GAKPred = parameters["ГАК председатель"];
                DPruk = parameters["ДП руководство"];
                DopuskVKR = parameters["Допуск к ВКР"];
                RetzVKR = parameters["Рецензия ВКР"];
                DPRetz = parameters["ДП рецензии"];
                MAGRuk = parameters["Магистранты руководство"];
                MagRetz = parameters["Магистранты рецензирование"];
                RukKaf = parameters["Руководство кафедрой"];
                NIIR = parameters["Научная работа"];
                NIIRRukMag = parameters["Руководство НИИР магистра"];
                ASPpractice = parameters["Практика аспирантов"];
                NIIRRukAsp = parameters["Руководство НИИР аспирантов"];
                DopuskDissMag = parameters["Допуск к защите диссертации магистрантов"];
                NormocontrolMag = parameters["Нормоконтроль  диссертации магистрантов"];
                DopuskBak = parameters["Допуск к защите бакалавров"];
                NormocontrolBak = parameters["Нормоконтроль бакалавров"];
                GosEkz = parameters["Государственный экзамен"];
                EkzBoard = parameters["Кол-во членов экз. комиссии"];
                cn.Close();
            }
            catch(Exception err)
            {
                MessageBox.Show("Не удалось загрузить настройки:\r\n" + err.Message + "\r\n" + err.StackTrace);
            }
        }
        static void UpdateDataBase(System.Data.OleDb.OleDbConnection cn, string column,float value)
        {
            string query = "UPDATE Normas SET [Value]=? WHERE ParameterName=?";
            var cmd = new System.Data.OleDb.OleDbCommand();
            cmd.Connection = cn;
            cmd.CommandType = System.Data.CommandType.Text;
            cmd.CommandText = query;
            cmd.Parameters.Add("@p1", System.Data.OleDb.OleDbType.Single).Value = value;
            cmd.Parameters.Add("@p2", System.Data.OleDb.OleDbType.Char).Value = column;
            cmd.ExecuteNonQuery();
        }
        public static void ToDataBase()
        {
            try
            {
                bool res = false;
                Model.DataManager.SharedDataManager();
                var cn = new System.Data.OleDb.OleDbConnection(Model.DataManager.Connection.ConnectionString);
                cn.Open();

                UpdateDataBase(cn, "Лекция", LectureCost);
                UpdateDataBase(cn, "Практическая работа", PracticeCost);
                UpdateDataBase(cn, "Лабораторная работа", LabCost);
                UpdateDataBase(cn, "Консультация", KonsCost);
                UpdateDataBase(cn, "Экзамен",(float)Math.Round((double)EkzCost,2));
                UpdateDataBase(cn, "КР", KRCost);
                UpdateDataBase(cn, "КП", KPCost);
                UpdateDataBase(cn, "Зачет", ZachCost);
                UpdateDataBase(cn, "РГР", RGR);
                UpdateDataBase(cn, "Учебная практика", UchPr);
                UpdateDataBase(cn, "Производственная практика", PrPr);
                UpdateDataBase(cn, "Преддипломная практика", PreddipPr);
                UpdateDataBase(cn, "Контрольные задания заочников", KzZaoch);
                UpdateDataBase(cn, "ГЭК", GEK);
                UpdateDataBase(cn, "ГАК", GAK);
                UpdateDataBase(cn, "ГАК председатель", GAKPred);
                UpdateDataBase(cn, "ДП руководство", DPruk);
                UpdateDataBase(cn, "Допуск к ВКР", DopuskVKR);
                UpdateDataBase(cn, "Рецензия ВКР", RetzVKR);
                UpdateDataBase(cn, "ДП рецензии", DPRetz);
                UpdateDataBase(cn, "Магистранты руководство", MAGRuk);
                UpdateDataBase(cn, "Магистранты рецензирование", MagRetz);
                UpdateDataBase(cn, "Руководство кафедрой", RukKaf);
                UpdateDataBase(cn, "Научная работа", NIIR);
                UpdateDataBase(cn, "Руководство НИИР магистра", NIIRRukMag);
                UpdateDataBase(cn, "Практика аспирантов", ASPpractice);
                UpdateDataBase(cn, "Руководство НИИР аспирантов", NIIRRukAsp);
                UpdateDataBase(cn, "Допуск к защите диссертации магистрантов", DopuskDissMag);
                UpdateDataBase(cn, "Нормоконтроль  диссертации магистрантов", NormocontrolMag);
                UpdateDataBase(cn, "Допуск к защите бакалавров", DopuskBak);
                UpdateDataBase(cn, "Нормоконтроль бакалавров", NormocontrolBak);
                UpdateDataBase(cn, "Государственный экзамен", GosEkz);
                UpdateDataBase(cn, "Кол-во членов экз. комиссии", EkzBoard);
                cn.Close();

            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }
        }
        public static void ToDefault()
        {
            try
            {
                RegistryKey saveKey = Registry.CurrentUser.CreateSubKey(@"software\ChairDB\CalculationSettings");
                saveKey.SetValue("LectureCost", 1f);
                saveKey.SetValue("LabCost", 1f);
                saveKey.SetValue("PracticeCost", 1f);
                saveKey.SetValue("KonsCost", 0.05f);
                saveKey.SetValue("EkzCost", 0.33f);
                saveKey.SetValue("KRCost", 2f);
                saveKey.SetValue("KPCost", 3f);
                saveKey.SetValue("ZachCost", 0.25f);

                saveKey.SetValue("RGR", 0.25f);
                saveKey.SetValue("UchPr", 4f);
                saveKey.SetValue("PrPr", 2f);
                saveKey.SetValue("PreddipPr", 1f);
                saveKey.SetValue("KzZaoch", 0.25f);
                saveKey.SetValue("GEK", 0.25f);
                saveKey.SetValue("GAK", 0.5f);
                saveKey.SetValue("GAKPred", 0.25f);
                saveKey.SetValue("DPruk", 14f);
                saveKey.SetValue("DopuskVKR", 0.25f);
                saveKey.SetValue("RetzVKR", 0.25f);
                saveKey.SetValue("DPRetz", 0.25f);
                saveKey.SetValue("MAGRuk", 36f);
                saveKey.SetValue("MagRetz", 0.25f);
                saveKey.SetValue("RukKaf", 60f);
                saveKey.SetValue("NIIR", 3f);
                saveKey.SetValue("NIIRRukMag", 0.25f);
                saveKey.SetValue("ASPpractice", 0.25f);
                saveKey.SetValue("NIIRRukAsp", 0.25f);
                saveKey.SetValue("DopuskDissMag", 1f);
                saveKey.SetValue("NormocontrolMag", 1f);
                saveKey.SetValue("DopuskBak", 0.5f);
                saveKey.SetValue("NormocontrolBak", 0.5f);
                saveKey.SetValue("AspRuk", 50f);
                saveKey.SetValue("GosEkz", 0.5f);
                saveKey.SetValue("EkzBoard", 6f);

                saveKey.Close();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
    }

    namespace ExportSettings
    {
        public static class SemesterSetting
        {
            public static int FITStartRow = 9;
            public static int MSFStartRow = 16;
            public static int IDPOStartRow = 24;
            public static int MAGStartRow = 32;
            public static int IndexColumn = 1;
            public static int GroupColumn = 2;
            public static int CourceColumn = 3;
            public static int DisciplineColumn = 4;
            public static int DisciplineCostColumn = 5;
            public static int LecFactColumn = 6;
            public static int StudentsColumn = 9;
            public static int ContractColumn = 10;

            public static void ToRegistry()
            {
                try
                {
                    RegistryKey saveKey = Registry.CurrentUser.CreateSubKey(@"software\ChairDB\Export\Semester");
                    saveKey.SetValue("FITStartRow", FITStartRow);
                    saveKey.SetValue("MSFStartRow", MSFStartRow);
                    saveKey.SetValue("IDPOStartRow", IDPOStartRow);
                    saveKey.SetValue("MAGStartRow", MAGStartRow);
                    saveKey.SetValue("IndexColumn", IndexColumn);
                    saveKey.SetValue("GroupColumn", GroupColumn);
                    saveKey.SetValue("CourceColumn", CourceColumn);
                    saveKey.SetValue("DisciplineColumn", DisciplineColumn);
                    saveKey.SetValue("DisciplineCostColumn", DisciplineCostColumn);
                    saveKey.SetValue("LecFactColumn", LecFactColumn);
                    saveKey.SetValue("StudentsColumn", StudentsColumn);
                    saveKey.SetValue("ContractColumn", ContractColumn);
                    saveKey.Close();
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

            public static void FromRegistry()
            {
                using (RegistryKey readKey = Registry.CurrentUser.OpenSubKey(@"software\ChairDB\Export\Semester"))
                {
                    try
                    {
                        if (readKey != null)
                        {
                            FITStartRow = Convert.ToInt32(readKey.GetValue("FITStartRow"));
                            MSFStartRow = Convert.ToInt32(readKey.GetValue("MSFStartRow"));
                            IDPOStartRow = Convert.ToInt32(readKey.GetValue("IDPOStartRow"));
                            MAGStartRow = Convert.ToInt32(readKey.GetValue("MAGStartRow"));
                            IndexColumn = Convert.ToInt32(readKey.GetValue("IndexColumn"));
                            GroupColumn = Convert.ToInt32(readKey.GetValue("GroupColumn"));
                            CourceColumn = Convert.ToInt32(readKey.GetValue("CourceColumn"));
                            DisciplineColumn = Convert.ToInt32(readKey.GetValue("DisciplineColumn"));
                            DisciplineCostColumn = Convert.ToInt32(readKey.GetValue("DisciplineCostColumn"));
                            LecFactColumn = Convert.ToInt32(readKey.GetValue("LecFactColumn"));
                            StudentsColumn = Convert.ToInt32(readKey.GetValue("StudentsColumn"));
                            ContractColumn = Convert.ToInt32(readKey.GetValue("ContractColumn"));
                            readKey.Close();
                        }
                    }
                    catch (Exception err)
                    {
                        MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            public static void ToDefault()
            {
                try
                {
                    RegistryKey saveKey = Registry.CurrentUser.CreateSubKey(@"software\ChairDB\Export\Semester");
                    saveKey.SetValue("FITStartRow", 9);
                    saveKey.SetValue("MSFStartRow", 16);
                    saveKey.SetValue("IDPOStartRow", 24);
                    saveKey.SetValue("MAGStartRow", 32);
                    saveKey.SetValue("IndexColumn", 1);
                    saveKey.SetValue("GroupColumn", 2);
                    saveKey.SetValue("CourceColumn", 3);
                    saveKey.SetValue("DisciplineColumn", 4);
                    saveKey.SetValue("DisciplineCostColumn", 5);
                    saveKey.SetValue("LecFactColumn", 6);
                    saveKey.SetValue("StudentsColumn", 9);
                    saveKey.SetValue("ContractColumn", 10);
                    saveKey.Close();
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public static class WorkloadSetting
        {

        }

        public static class IndPlanSetting
        {
            public static int DisciplinesStartRow = 5;
            public static int[] YearsDescriptionRowCell = { 26, 6 };
            public static int[] TeacherFIORowCell = { 29, 8 };
            public static int DisciplineNameColumn = 3;
            public static int DisciplineDescrColumn = 4;
            public static int StudentsCountColumn = 5;
            public static int DisciplineSettingsStartColumn = 9;

            public static void ToRegistry()
            {
                try
                {
                    RegistryKey saveKey = Registry.CurrentUser.CreateSubKey("software\\ChairDB\\Export\\IndPlan");
                    saveKey.SetValue("DisciplinesStartRow", DisciplinesStartRow);
                    saveKey.SetValue("YearsDescriptionRowCell", YearsDescriptionRowCell[0] + "/" + YearsDescriptionRowCell[1]);
                    saveKey.SetValue("TeacherFIORowCell", TeacherFIORowCell[0] + "/" + TeacherFIORowCell[1]);
                    saveKey.SetValue("DisciplineNameColumn", DisciplineNameColumn);
                    saveKey.SetValue("DisciplineDescrColumn", DisciplineDescrColumn);
                    saveKey.SetValue("StudentsCountColumn", StudentsCountColumn);
                    saveKey.SetValue("DisciplineSettingsStartColumn", DisciplineSettingsStartColumn);
                    saveKey.Close();
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void FromRegistry()
            {
                try
                {
                    RegistryKey readKey = Registry.CurrentUser.OpenSubKey("software\\ChairDB\\Export\\IndPlan");
                    if(readKey!=null)
                    {
                        DisciplinesStartRow = Convert.ToInt32(readKey.GetValue("DisciplinesStartRow"));

                        var buf = readKey.GetValue("YearsDescriptionRowCell").ToString().Split('/');
                        YearsDescriptionRowCell =new int[]{ Convert.ToInt32(buf[0]), Convert.ToInt32(buf[1])};

                        buf = readKey.GetValue("TeacherFIORowCell").ToString().Split('/');
                        TeacherFIORowCell = new int[] { Convert.ToInt32(buf[0]), Convert.ToInt32(buf[1]) };

                        DisciplineNameColumn = Convert.ToInt32(readKey.GetValue("DisciplineNameColumn"));
                        DisciplineDescrColumn = Convert.ToInt32(readKey.GetValue("DisciplineDescrColumn"));
                        StudentsCountColumn = Convert.ToInt32(readKey.GetValue("StudentsCountColumn"));
                        DisciplineSettingsStartColumn = Convert.ToInt32(readKey.GetValue("DisciplineSettingsStartColumn"));
                        readKey.Close();
                    }   
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public static class ContractSetting
        {
            public static string TeacherName = "Teacher name";
            public static string DisciplineTableTag = "Discipline table";
            public static string RowNumberTag = "row number";
            public static string DisciplineTag = "Discipline";
            public static string GroupTag = "Group";
            public static string LectureTag = "Lec";
            public static string PractTag = "Pract";
            public static string LabTag = "Lab";
            public static string KrTag = "KR";
            public static string KpTag = "KP";
            public static string NirTag = "NIR";
            public static string RukPractTag = "RukPr";
            public static string ConsTag = "Cons";
            public static string ZachTag = "Zach";
            public static string EkzTag = "Ekz";
            public static string VkrTag = "VKR";
            public static string GekTag = "GEK";
            public static string RecVkrTag = "RecVKR";
            public static string TotalTag = "Total";
            public static string TotalTimeTag = "Total time";

            public static void ToRegistry()
            {
                try
                {
                    RegistryKey saveKey = Registry.CurrentUser.CreateSubKey("software\\ChairDB\\Export\\Contract");
                    saveKey.SetValue("TeacherName", TeacherName);
                    saveKey.SetValue("DisciplineTableTag", DisciplineTableTag);
                    saveKey.SetValue("RowNumberTag", RowNumberTag);
                    saveKey.SetValue("DisciplineTag", DisciplineTag);
                    saveKey.SetValue("GroupTag", TeacherName);
                    saveKey.SetValue("LectureTag", LectureTag);
                    saveKey.SetValue("PractTag", PractTag);
                    saveKey.SetValue("LabTag", LabTag);
                    saveKey.SetValue("KrTag", KrTag);
                    saveKey.SetValue("KpTag", KpTag);
                    saveKey.SetValue("NirTag", NirTag);
                    saveKey.SetValue("RukPractTag", RukPractTag);
                    saveKey.SetValue("ConsTag", ConsTag);
                    saveKey.SetValue("ZachTag", ZachTag);
                    saveKey.SetValue("EkzTag", EkzTag);
                    saveKey.SetValue("VkrTag", VkrTag);
                    saveKey.SetValue("RecVkrTag", RecVkrTag);
                    saveKey.SetValue("TotalTag", TotalTag);
                    saveKey.SetValue("TotalTimeTag", TotalTimeTag);
                    saveKey.Close();
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void FromRegistry()
            {
                try
                {
                    RegistryKey readKey = Registry.CurrentUser.OpenSubKey("software\\ChairDB\\Export\\IndPlan");
                    if (readKey != null)
                    {
                        TeacherName = readKey.GetValue("Teacher name").ToString();
                        DisciplineTableTag = readKey.GetValue("Discipline table").ToString();
                        RowNumberTag = readKey.GetValue("row number").ToString();
                        DisciplineTag = readKey.GetValue("Discipline").ToString();
                        GroupTag = readKey.GetValue("Group").ToString();
                        LectureTag = readKey.GetValue("Lec").ToString();
                        PractTag = readKey.GetValue("Pract").ToString();
                        LabTag = readKey.GetValue("Lab").ToString();
                        KrTag = readKey.GetValue("KR").ToString();
                        KpTag = readKey.GetValue("KP").ToString();
                        NirTag = readKey.GetValue("NIR").ToString();
                        RukPractTag = readKey.GetValue("RukPr").ToString();
                        ConsTag = readKey.GetValue("Cons").ToString();
                        ZachTag = readKey.GetValue("Zach").ToString();
                        EkzTag = readKey.GetValue("Ekz").ToString();
                        VkrTag = readKey.GetValue("VKR").ToString();
                        GekTag = readKey.GetValue("GEK").ToString();
                        RecVkrTag = readKey.GetValue("RecVKR").ToString();
                        TotalTag = readKey.GetValue("Total").ToString();
                        TotalTimeTag = readKey.GetValue("Total time").ToString();
                    }
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void ToDefault()
            {

            }

            public static string[] GetPrefixTag()
            {
                return new string[] { LectureTag, PractTag, LabTag, KrTag, KpTag, NirTag, RukPractTag,
                    ConsTag, ZachTag, EkzTag,VkrTag, GekTag, RecVkrTag, TotalTag};
            }
        }
    }
}
