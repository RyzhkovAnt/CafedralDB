using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApplicationSettings
{
	public static class ImportSettings
	{
		public static int StartReadingRow = 5;
		public static int GroupColumn = 1;
		public static int SemesterColumn = 2;
		public static int WeeksColumn = 3;
		public static int DisciplineNameColumn = 4;
		public static int LecturesColumn = 5;
		public static int PracticesColumn = 6;
		public static int LabsColumn = 7;
		public static int KzColumn = 8;
		public static int KrColumn = 9;
		public static int KpColumn = 10;
		public static int EkzColumn = 11;
		public static int ZachColumn = 12;
		public static int OtherColumn = 13;

		public static void ToRegistry()
		{
			RegistryKey saveKey = Registry.LocalMachine.CreateSubKey("software\\ChairDB\\Import");
			saveKey.SetValue("StartReadingRow", StartReadingRow);
			saveKey.SetValue("GroupColumn", GroupColumn);
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
			saveKey.Close();
		}

		public static void FromRegistry()
		{
			try
			{
				RegistryKey readKey = Registry.LocalMachine.OpenSubKey("software\\ChairDB\\Import");
				StartReadingRow = Convert.ToInt32(readKey.GetValue("StartReadingRow"));
				GroupColumn = Convert.ToInt32(readKey.GetValue("GroupColumn"));
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
				readKey.Close();
			}
			catch (Exception err)
			{ }
		}

	}

	public static class CalculationSettings
	{
		public static float LectureCost = 1;
		public static float LabCost = 1;
		public static float PracticeCost = 1;
		public static float KonsCost = 2;
		public static float EkzCost = 0.33f;
		public static float KRCost = 2;
		public static float KPCost = 3;
		public static float ZachCost = 0.25f;

		public static void ToRegistry()
		{
			RegistryKey saveKey = Registry.LocalMachine.CreateSubKey("software\\ChairDB\\CalculationSettings");
			saveKey.SetValue("LectureCost", LectureCost);
			saveKey.SetValue("LabCost", LabCost);
			saveKey.SetValue("PracticeCost", PracticeCost);
			saveKey.SetValue("KonsCost", KonsCost);
			saveKey.SetValue("EkzCost", EkzCost);
			saveKey.SetValue("KRCost", KRCost);
			saveKey.SetValue("KPCost", KPCost);
			saveKey.SetValue("ZachCost", ZachCost);
			saveKey.Close();
		}

		public static void FromRegistry()
		{
			try
			{
				RegistryKey readKey = Registry.LocalMachine.OpenSubKey("software\\ChairDB\\CalculationSettings");
				LectureCost = Convert.ToInt32(readKey.GetValue("LectureCost"));
				LabCost = Convert.ToInt32(readKey.GetValue("LabCost"));
				PracticeCost = Convert.ToInt32(readKey.GetValue("PracticeCost"));
				KonsCost = Convert.ToInt32(readKey.GetValue("KonsCost"));
				EkzCost = Convert.ToInt32(readKey.GetValue("EkzCost"));
				KRCost = Convert.ToInt32(readKey.GetValue("KRCost"));
				KPCost = Convert.ToInt32(readKey.GetValue("KPCost"));
				ZachCost = Convert.ToInt32(readKey.GetValue("ZachCost"));
				readKey.Close();
			}
			catch (Exception err)
			{ }
		}
	}
	namespace ExportSettings
	{
		public static class Semester
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

			public static void ToRegistry()
			{
				RegistryKey saveKey = Registry.LocalMachine.CreateSubKey("software\\ChairDB\\Export\\Semester");
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
				saveKey.Close();

			}

			public static void FromRegistry()
			{
				try
				{
					RegistryKey readKey = Registry.LocalMachine.OpenSubKey("software\\ChairDB\\Export\\Semester");
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
					readKey.Close();
				}
				catch (Exception err)
				{ }
			}
		}

		public static class Workload
		{

		}

		public static class IndPlan
		{
			public static int DisciplinesStartRow = 5;
			public static int[] YearsDescriptionRowCell = {26,6};
			public static int[] TeacherFIORowCell = { 29, 8 };
			public static int DisciplineNameColumn = 3;
			public static int DisciplineDescrColumn = 4;
			public static int StudentsCountColumn =  5;
			public static int DisciplineSettingsStartColumn = 9;

			public static void ToRegistry()
			{
				RegistryKey saveKey = Registry.LocalMachine.CreateSubKey("software\\ChairDB\\Export\\IndPlan");

				saveKey.Close();

			}

			public static void FromRegistry()
			{
				try
				{
					RegistryKey readKey = Registry.LocalMachine.OpenSubKey("software\\ChairDB\\Export\\IndPlan");

					readKey.Close();
				}
				catch (Exception err)
				{ }
			}
		}


	}


}
