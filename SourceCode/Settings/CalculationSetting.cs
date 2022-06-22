using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Windows.Forms;
using CafedralDB.SourceCode.Model;
namespace CafedralDB.SourceCode.Settings
{
    internal class CalculationSetting : ISetting
    {

        public float LectureCost { get; private set; } = 1;
        public float LabCost { get; private set; } = 1;
        public float PracticeCost { get; private set; } = 1;
        public float KonsCost { get; private set; } = 0.05f;
        public float EkzCost { get; private set; } = 0.33f;
        public float KRCost { get; private set; } = 2;
        public float KPCost { get; private set; } = 3;
        public float ZachCost { get; private set; } = 0.25f;

        public float RGR { get; private set; } = 0.25f;
        public float UchPr { get; private set; } = 4;
        public float PrPr { get; private set; } = 2;
        public float PreddipPr { get; private set; } = 1;
        public float KzZaoch { get; private set; } = 0.25f;
        public float GEK { get; private set; } = 0.5f;
        public float GAK { get; private set; } = 0.5f;
        public float GAKPred { get; private set; } = 0.25f;
        public float DPruk { get; private set; } = 14;
        public float DopuskVKR { get; private set; } = 0.5f;
        public float RetzVKR { get; private set; } = 0.25f;
        public float DPRetz { get; private set; } = 0.25f;
        public float MAGRuk { get; private set; } = 36;
        public float MagRetz { get; private set; } = 0.25f;
        public float RukKaf { get; private set; } = 60;
        public float NIIR { get; private set; } = 3.0f;
        public float NIIRRukMag { get; private set; } = 0.25f;
        public float ASPpractice { get; private set; } = 0.25f;
        public float NIIRRukAsp { get; private set; } = 0.25f;
        public float DopuskDissMag { get; private set; } = 1;
        public float NormocontrolMag { get; private set; } = 1;
        public float DopuskBak { get; private set; } = 0.5f;
        public float NormocontrolBak { get; private set; } = 0.5f;
        public float AspRuk { get; private set; } = 50f;
        public float GosEkz { get; private set; } = 0.5f;
        public float EkzBoard { get; private set; } = 6;
        public CalculationSetting()
        {
            LoadFromRegistry();
        }

        public void LoadFromRegistry()
        {
            try
            {
                RegistryKey readkey = Registry.CurrentUser.OpenSubKey("software\\ChairDB\\CalculationSettings");
                if (readkey != null)
                {
                    var curProps = this.GetType().GetProperties();
                    var SettingFields = typeof(CalculationSettingFields).GetFields();
                    foreach (var field in SettingFields)
                    {
                        for (int i = 0; i < curProps.Length; i++)
                        {
                            if (curProps[i].Name == field.Name)
                            {
                                var newValue = Convert.ToSingle(readkey.GetValue(field.Name));
                                curProps[i].SetValue(this, newValue);
                                break;
                            }
                        }
                    }
                    readkey.Close();
                }
                else
                {
                    ReadFromDataBase();
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        public void ToDefault()
        {
            try
            {
                Type type = typeof(CalculationSettingFields);
                var fieldInfos = type.GetFields();
                RegistryKey savekey = Registry.CurrentUser.CreateSubKey("software\\ChairDB\\CalculationSettings");
                DataManager.SharedDataManager();
                var cn = new System.Data.OleDb.OleDbConnection(DataManager.Connection.ConnectionString);
                cn.Open();
                foreach (var field in fieldInfos)
                {
                    var tmp = (Field)field.GetValue(null);
                    savekey.SetValue(tmp.Name, tmp.DefaultValue.ToString());
                    SaveToDataBase(cn, tmp.DataBaseName, Convert.ToSingle(tmp.DefaultValue));
                }
                savekey.Close();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void SaveCurrentInRegistry()
        {
            try
            {
                Type type = typeof(CalculationSettingFields);
                var fieldInfos = type.GetFields();
                var curProps = this.GetType().GetProperties();
                RegistryKey savekey = Registry.CurrentUser.CreateSubKey("software\\ChairDB\\CalculationSettings");
                foreach (var field in fieldInfos)
                {
                    for (int i = 0; i < curProps.Length; i++)
                    {
                        if (curProps[i].Name == field.Name)
                        {
                            var tmp = (float)curProps[i].GetValue(this);
                            var fieldName = (Field)field.GetValue(null);
                            savekey.SetValue(fieldName.Name, tmp.ToString());

                        }
                    } 
                
                }
                savekey.Close();

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
                RegistryKey saveKey = Registry.CurrentUser.CreateSubKey("software\\ChairDB\\CalculationSettings");
                DataManager.SharedDataManager();
                var cn = new System.Data.OleDb.OleDbConnection(DataManager.Connection.ConnectionString);
                cn.Open();
                foreach (Field key in newValues.Keys)
                {
                    saveKey.SetValue(key.Name, newValues[key]);
                    SaveToDataBase(cn, key.DataBaseName, (float)newValues[key]);
                }
                saveKey.Close();
                cn.Close();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Ошибка при сохранении настроек импорта", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ReadFromDataBase()
        {
            try
            {
                DataManager.SharedDataManager();
                var cn = new System.Data.OleDb.OleDbConnection(Model.DataManager.Connection.ConnectionString);
                var cmd = new System.Data.OleDb.OleDbCommand();
                cn.Open();
                cmd.Connection = cn;
                cmd.CommandType = System.Data.CommandType.Text;
                cmd.CommandText = "Select ParameterName, Value From Normas";
                System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
                Dictionary<string, float> parameters = new Dictionary<string, float>();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        parameters.Add(reader[0].ToString(), Convert.ToSingle(reader[1].ToString()));
                    }
                }
                cn.Close();
                if (parameters.Count > 0)
                {
                    var curProps = this.GetType().GetProperties();
                    var SettingFields = typeof(CalculationSettingFields).GetFields();

                    foreach (var field in SettingFields)
                    {
                        for (int i = 0; i < curProps.Length; i++)
                        {
                            if (field.Name == curProps[i].Name)
                            {
                                var tmp = (Field)field.GetValue(null);
                                curProps[i].SetValue(this, parameters[tmp.DataBaseName]);
                            }
                        }
                    }
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Ошибка при сохранении настроек импорта", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        void SaveToDataBase(System.Data.OleDb.OleDbConnection cn, string column, float value)
        {
            string query = "UPDATE Normas SET [Value]=? WHERE ParameterName=?";
            var cmd = new System.Data.OleDb.OleDbCommand();
            cmd.Connection = cn;
            cmd.CommandType = System.Data.CommandType.Text;
            cmd.CommandText = query;
            cmd.Parameters.Add("@p1", System.Data.OleDb.OleDbType.Single).Value = Math.Round(value,2);
            cmd.Parameters.Add("@p2", System.Data.OleDb.OleDbType.Char).Value = column;
            cmd.ExecuteNonQuery();
        }
    }
}
