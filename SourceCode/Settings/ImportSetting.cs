using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Win32;

namespace CafedralDB.SourceCode.Settings
{
    internal class ImportSetting:ISetting
    {
        public int StartReadingRow { get; private set; } = 7;
        public int GroupColumn { get; private set; } = 1;
        public int GroupCountColumn { get; private set; } = 5;
        public int StudentCountColumn { get; private set; } = 2;
        public int SemesterColumn { get; private set; } = 6;
        public int WeeksColumn { get; private set; } = 7;
        public int DisciplineNameColumn { get; private set; } = 8;
        public int LecturesColumn { get; private set; } = 10;
        public int PracticesColumn { get; private set; } = 11;
        public int LabsColumn { get; private set; } = 12;
        public int KzColumn { get; private set; } = 17;
        public int KrColumn { get; private set; } = 13;
        public int KpColumn { get; private set; } = 14;
        public int EkzColumn { get; private set; } = 15;
        public int ZachColumn { get; private set; } = 16;
        public int SpecialColumn { get; private set; } = 18;

        public ImportSetting()
        {
            LoadFromRegistry();
        }
        public void ToDefault()
        {
            try
            {
                Type type = typeof(ImportSettingFields);
                System.Reflection.FieldInfo[] fieldInfos = type.GetFields();

                RegistryKey saveKey = Registry.CurrentUser.CreateSubKey("software\\ChairDB\\Import");
                foreach (var field in fieldInfos)
                {
                    var tmp = (Field)field.GetValue(null);
                    saveKey.SetValue(tmp.Name, tmp.DefaultValue.ToString());

                }
                saveKey.Close();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void LoadFromRegistry()
        {
            try
            {
                RegistryKey readKey = Registry.CurrentUser.OpenSubKey(@"software\ChairDB\Import");
                if (readKey != null)
                {
                    var tmpFields = this.GetType().GetProperties();
                    var inputType = typeof(ImportSettingFields);
                    foreach (var field in inputType.GetFields())
                    {
                        for (int i = 0; i < tmpFields.Length; i++)
                        {
                            if (tmpFields[i].Name == field.Name)
                            {
                                var newValue = Convert.ToInt32(readKey.GetValue(field.Name));
                                tmpFields[i].SetValue(this, newValue);
                                break;
                            }
                        }
                        //если дошли до конца то параметра нет
                    }
                    readKey.Close();
                }
                else
                {
                    var tmpFields = this.GetType().GetProperties();
                    var inputType = typeof(ImportSettingFields);
                    foreach (var field in inputType.GetFields())
                    {
                        for (int i = 0; i < tmpFields.Length; i++)
                        {
                            if (tmpFields[i].Name == field.Name)
                            {
                                var newValue = (Field)field.GetValue(null);
                                tmpFields[i].SetValue(this, newValue.DefaultValue);
                                break;
                            }
                        }
                        //если дошли до конца то параметра нет
                    }
                }

            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void SaveToRegistry(Dictionary<Field, object> newValues)
        {
            try
            {
                RegistryKey saveKey = Registry.CurrentUser.CreateSubKey("software\\ChairDB\\Import");
                foreach (var key in newValues.Keys)
                {
                    saveKey.SetValue(key.Name, newValues[key]);
                }
                saveKey.Close();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Ошибка при сохранении настроек импорта", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
