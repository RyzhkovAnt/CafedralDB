using CafedralDB;
using CafedralDB.MainDBDataSetTableAdapters;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using static CafedralDB.MainDBDataSet;
using CafedralDB.SourceCode.Model.Entity;

namespace CafedralDB.SourceCode.Model
{
	public class DataManager
	{
		static string databasePath = "\\Database\\MainDB.accdb";
		public static string dbPath = "";
		static string connString = "";
		static List<StudyYear> years = new List<StudyYear>();
		static List<Semester> semesters = new List<Semester>();
		static OleDbConnection cn;
		#region singletone
		static DataManager _instance;

		public static OleDbConnection Connection
		{
			get
			{
				return cn;
			}
		}

		public DataManager()
		{
			dbPath = System.Windows.Forms.Application.StartupPath + databasePath;
			connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DataManager.dbPath + ";" +
										"Persist Security Info = False; ";
			cn = new System.Data.OleDb.OleDbConnection(connString);
		}

		public static DataManager SharedDataManager()
		{
			if (_instance == null)
				_instance = new DataManager();

			return _instance;
		}

		public void CloseConnection()
		{
			cn.Close();
		}

		public void OpenConnection()
		{
			if (cn.State != ConnectionState.Open)
				cn.Open();
		}

		#endregion

		#region Get code
		public static List<Semester> InitSemestersFromDB()
		{
			semesters.Clear();
			//MainDBDataSet mainDBDataset = new MainDBDataSet();
			//mainDBDataset.DataSetName = "MainDBDataSet";
			//mainDBDataset.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
			//((System.ComponentModel.ISupportInitialize)(mainDBDataset)).BeginInit();

			//SemestersTableAdapter semesterTableAdapter = new SemestersTableAdapter();
			//semesterTableAdapter.ClearBeforeFill = true;
			//semesterTableAdapter.Fill(mainDBDataset.Semesters);

			//((System.ComponentModel.ISupportInitialize)(mainDBDataset)).EndInit();

			foreach (DataRow row in DataSetHolder.GetInstance().SemestersQuery)
			{
				Semester semester = new Semester((int)row["ID"]);
				semester.Name = (string)row["Descr"];
				semester.WeekCount = (int)row["WeekCount"];
				semesters.Add(semester);
			}

			return semesters;
		}

		public int GetSemesterIDByName(string name)
		{
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Semester Where Descr = @Descr";
			cmd.Parameters.AddWithValue("@Descr", name);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();
			int result;
			if (reader.HasRows)
				result = Convert.ToInt32(reader[0]);
			else
				result = -1;

			cn.Close();
			return result;
		}

		public int GetSpecialityIDByName(string name)
		{
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Speciality Where Descr = @Descr";
			cmd.Parameters.AddWithValue("@Descr", name);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			int result;
			if (reader.HasRows)
				result = Convert.ToInt32(reader[0]);
			else
				result = -1;

			cn.Close();
			return result;
		}

		public int FindTypeByName(string disciplineName)
		{
			int result = 1;

			var cn = new OleDbConnection(connString);
			var cmd = new OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From DisciplineType Where Descr = @Descr";
			cmd.Parameters.AddWithValue("@Descr", disciplineName);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			if (reader.HasRows)
				result = Convert.ToInt32(reader[0]);
			cn.Close();


			return result;
		}

		public int GetSpecialityIDByGroupID(int GroupID)
		{
			var cmd = CreateQuery("SELECT        Speciality.ID AS Expr1 " +
						"FROM([Group] INNER JOIN " +
						 "Speciality ON[Group].Speciality = Speciality.ID) " +
						"WHERE([Group].ID = @param)");
			cmd.Parameters.AddWithValue("@param", GroupID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			int result;
			if (reader.HasRows)
				result = Convert.ToInt32(reader[0]);
			else
				result = -1;
			CloseConnection();
			return result;
		}

		public int GetYearIDByName(string year)
		{
			var cmd = CreateQuery("Select * From StudyYear Where StudyYear = @Descr");
			cmd.Parameters.AddWithValue("@Descr", Convert.ToInt32(year));
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			int result;
			if (reader.HasRows)
				result = Convert.ToInt32(reader[0]);
			else
				result = -1;
			CloseConnection();
			return result;
		}

		public int GetGroupIDByYearAndSpeciality(int year, int spec)
		{
			var cmd = CreateQuery("SELECT [Group].ID, StudyYear.ID AS yearID, Speciality.ID AS specID " +
									"FROM(([Group] INNER JOIN " +
						 "StudyYear ON[Group].EntryYear = StudyYear.ID) INNER JOIN " +
						 "Speciality ON[Group].Speciality = Speciality.ID) " +
						"WHERE(StudyYear.ID = @year) AND(Speciality.ID = @spec)");
			cmd.Parameters.AddWithValue("@year", year);
			cmd.Parameters.AddWithValue("@spec", spec);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			int result;
			if (reader.HasRows)
				result = Convert.ToInt32(reader[0]);
			else
				result = -1;
			CloseConnection();
			return result;
		}

		public int FindLastWorkloadAssign(Workload workload, string DisciplineDescr, int specialityID)
		{
			int EmployeeID = -1;

			var cmd = CreateQuery("SELECT WorkloadAssign.Teacher " +
						"FROM            ((((WorkloadAssign INNER JOIN " +
"Workload ON WorkloadAssign.Workload = Workload.ID) INNER JOIN " +
"Discipline ON Workload.Discipline = Discipline.ID) INNER JOIN " +
"[Group] ON Workload.[Group] = [Group].ID) INNER JOIN " +
"Speciality ON[Group].Speciality = Speciality.ID) " +
"WHERE(Discipline.Descr = @param1) AND(Workload.Semester = @param2) AND(Speciality.ID = @param3) AND(Workload.StudyYear = @param4)");
			cmd.Parameters.AddWithValue("@param1", DisciplineDescr);
			cmd.Parameters.AddWithValue("@param2", workload.Semester);
			cmd.Parameters.AddWithValue("@param3", specialityID);
			cmd.Parameters.AddWithValue("@param4", workload.Year - 1);

			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			if (reader.HasRows)
				EmployeeID = Convert.ToInt32(reader[0]);
			else
				EmployeeID = -1;
			CloseConnection();

			return EmployeeID;
		}
		#endregion

		public void ClearYear(string year)
		{
			List<int> workloadsIDs = new List<int>();
			List<int> workloadsAssignsIDs = new List<int>();
			HashSet<int> disciplines = new HashSet<int>();

			int yearID = GetYearIDByName(year);

			var cmd = CreateQuery("Select * From Workload Where StudyYear = @YearID");
			cmd.Parameters.AddWithValue("@YearID", yearID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			while (reader.Read())
			{
				workloadsIDs.Add(Convert.ToInt32(reader[0]));

				var cmd2 = CreateQuery("DELETE FROM WorkloadAssign WHERE(WorkloadAssign.Workload = @WorkloadID)");
				cmd2.Parameters.AddWithValue("@WorkloadID", reader[0]);
				cmd2.ExecuteNonQuery();
			}

			foreach (int workload in workloadsIDs)
			{
				var cmd2 = CreateQuery("DELETE FROM Workload WHERE(ID = @WorkloadID)");
				cmd2.Parameters.AddWithValue("@WorkloadID", workload);
				cmd2.ExecuteNonQuery();

				cmd2 = CreateQuery("DELETE FROM Discipline WHERE(ID = @DisciplineID)");
				cmd2.Parameters.AddWithValue("@DisciplineID", workload);
				cmd2.ExecuteNonQuery();
			}
		}

		public OleDbCommand CreateQuery(string query)
		{
			var cmd = new System.Data.OleDb.OleDbCommand();
			if (cn.State != ConnectionState.Open)
				cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;

			cmd.CommandText = query;

			return cmd;
		}

		#region GetData

		#region ByID
		public AcademicDegree GetAcademicDegree(int ID)
		{
			AcademicDegree result;
			var cmd = CreateQuery("Select * From AcademicDegree Where ID = @ID");
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new AcademicDegree(Convert.ToInt32(reader[0]));
			result.Name = reader[1].ToString();
			CloseConnection();
			return result;
		}

		public AcademicRank GetAcademicRank(int ID)
		{
			AcademicRank result;

			var cmd = CreateQuery("Select * From AcademicRank Where ID = @ID");
			cmd.CommandText = "Select * From AcademicRank Where ID = @ID";
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new AcademicRank(Convert.ToInt32(reader[0]));
			result.Name = reader[1].ToString();
			CloseConnection();
			return result;
		}

		public Department GetDepartment(int ID)
		{
			Department result;
			var cmd = CreateQuery("Select * From Department Where ID = @ID");
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new Department(Convert.ToInt32(reader[0]));
			result.CodeNumber = Convert.ToInt32(reader[1]);
			result.Description = reader[1].ToString();
			CloseConnection();
			return result;
		}

		public Discipline GetDiscipline(int ID)
		{
			Discipline result;

			var cmd = CreateQuery("Select * From Discipline Where ID = @ID");
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			/*
			 *	0 -ID
				1 - "Descr"
				2 - "Department"
				3 -"DisciplineType"
				4 -"LectureCount"
				5 -"PracticeCount"
				6 -"LabCount"
				7 -"KR"
				8 -"KP"
				9 -"RGR"
				10 -"Zach"
				11 -"Ekz"
				12 -"Kons"
				13 -"UchPr"
				14 -"prPr"
				15 -"predDipPr"
				16 -"NIIR"
				17 -"KonsZaoch"
				18 -"GEK"
				19 -"GAK"
				20 -"GosEkz"
				21 -"GAKPred"
				22 -"DPruk"
				23 -"DopuskVKR"
				24 -"RetzVKR"
				25 -"DPretz"
				26 -"ASPRuk"
				27 -"MAGRuk"
				28 -"MAGRetz"
				29 -"RukKaf"
				30 -"isSpecial"
				31 -"Contr"

			 */

			result = new Discipline(Convert.ToInt32(reader[0]));
			result.Descr = reader[1].ToString();
			result.DepartmentID = Convert.ToInt32(reader[2]);
			result.TypeID = Convert.ToInt32(reader[3]);
			result.LectureCount = Convert.ToInt32(reader[4]);
			result.PracticeCount = Convert.ToInt32(reader[5]);
			result.LabCount = Convert.ToInt32(reader[6]);
			result.KR = Convert.ToBoolean(reader[7]);
			result.KP = Convert.ToBoolean(reader[8]);
			result.RGR = Convert.ToBoolean(reader[9]);
			result.Zach = Convert.ToBoolean(reader[10]);
			result.Ekz = Convert.ToBoolean(reader[11]);
			result.Kons = Convert.ToBoolean(reader[12]);
			result.UchPr = Convert.ToInt32(reader[13]);
			result.PrPr = Convert.ToInt32(reader[14]);
			result.PredDipPr = Convert.ToInt32(reader[15]);
			result.NIIR = Convert.ToInt32(reader[16]);
			result.KonsZaoch = Convert.ToBoolean(reader[17]);
			result.GEK = Convert.ToBoolean(reader[18]);
			result.GAK = Convert.ToBoolean(reader[19]);
			result.GosEkz = Convert.ToBoolean(reader[20]);
			result.GAKPred = Convert.ToBoolean(reader[21]);
			result.DPRuk = Convert.ToBoolean(reader[22]);
			result.DopuskVKR = Convert.ToBoolean(reader[23]);
			result.RetzVKR = Convert.ToBoolean(reader[24]);
			result.DPretz = Convert.ToBoolean(reader[25]);
			result.ASPRuk = Convert.ToBoolean(reader[26]);
			result.MAGRuk = Convert.ToBoolean(reader[27]);
			result.MAGRetz = Convert.ToBoolean(reader[28]);
			result.RukKaf = Convert.ToBoolean(reader[29]);
			if (reader[30] != DBNull.Value)
				result.Special = Convert.ToBoolean(reader[30]);
			CloseConnection();
			return result;
		}

		public Employee GetEmployee(int ID)
		{
			Employee result;
			//var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Employee Where ID = @ID";
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new Employee(Convert.ToInt32(reader[0]));
			result.Name = reader[1].ToString();
			if (reader[2] != System.DBNull.Value)
				result.ContractNumber = Convert.ToInt32(reader[2]);
			if (reader[3] != System.DBNull.Value)
				result.Rate = Convert.ToSingle(reader[3]);
			if (reader[4] != System.DBNull.Value)
				result.BirthDay = Convert.ToDateTime(reader[4]);
			if (reader[5] != System.DBNull.Value)
				result.WorkingPositionID = Convert.ToInt32(reader[5]);
			if (reader[6] != System.DBNull.Value)
				result.AcademicRankID = Convert.ToInt32(reader[6]);
			if (reader[7] != System.DBNull.Value)
				result.AcademicDegreeID = Convert.ToInt32(reader[7]);
			CloseConnection();
			return result;
		}

		public Faculty GetFaculty(int ID)
		{
			Faculty result;
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Faculty Where ID = @ID";
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new Faculty(Convert.ToInt32(reader[0]));
			result.Name = reader[1].ToString();
			result.FullName = reader[2].ToString();
			CloseConnection();
			return result;
		}

		public Group GetGroup(int ID)
		{
			Group result;
			//var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From [Group] Where ID = @ID";
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new Group(Convert.ToInt32(reader[0]));
			result.Name = reader[1].ToString();
			result.Speciality = Convert.ToInt32(reader[2]);
			result.StudyQualification = Convert.ToInt32(reader[3]);
			result.StudyFormID = Convert.ToInt32(reader[4]);
			result.StudentCount = Convert.ToInt32(reader[5]);
			result.EntryYear = Convert.ToInt32(reader[6]);
            result.SubgroupCount = Convert.ToInt32(reader[7]);
			CloseConnection();
			return result;
		}

		public Qualification GetQualification(int ID)
		{
			Qualification result;
			//var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Qualification Where ID = @ID";
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new Qualification(Convert.ToInt32(reader[0]));
			result.Name = reader[1].ToString();
			CloseConnection();
			return result;
		}

		public Semester GetSemester(int ID)
		{
			Semester result;
			//var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Semester Where ID = @ID";
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new Semester(Convert.ToInt32(reader[0]));
			result.Name = reader[1].ToString();
			result.WeekCount = Convert.ToInt32(reader[2]);
			CloseConnection();
			return result;
		}

		public Speciality GetSpeciality(int ID)
		{
			Speciality result;
			//var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Speciality Where ID = @ID";
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new Speciality(Convert.ToInt32(reader[0]));
			result.Name = reader[1].ToString();
			result.Faculty = Convert.ToInt32(reader[2]);
			CloseConnection();
			return result;
		}

		public Student GetStudent(int ID)
		{
			Student result;
			//var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Student Where ID = @ID";
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new Student(Convert.ToInt32(reader[0]));
			result.Name = reader[1].ToString();
			CloseConnection();
			return result;
		}

		public StudyYear GetStudyYear(int ID)
		{
			StudyYear year;
			//var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			OpenConnection();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From StudyYear Where ID = @ID";
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			reader.Read();
			// write the data on to the screen
			if (reader.HasRows)
			{
				Console.WriteLine(String.Format("{0} \t | {1} \t",
				// call the objects from their index
				reader[0], reader[1]));

				//DataRow query = DataSetHolder.GetInstance().SemestersQuery;
				//foreach (DataRow row in DataSetHolder.GetInstance().SemestersQuery)
				year = new StudyYear(Convert.ToInt32(reader[0]));
				year.Year = Convert.ToInt32(reader[1]);
				//	

				CloseConnection();
				return year;
			}
			CloseConnection();
			return null;
		}

		public Workload GetWorkload(int ID)
		{
			Workload result;
			//var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			if (cn.State != ConnectionState.Open)
				cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Workload Where ID = @ID";
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new Workload(Convert.ToInt32(reader[0]));
			result.Discipline = Convert.ToInt32(reader[1]);
			result.Year = Convert.ToInt32(reader[2]);
			result.Semester = Convert.ToInt32(reader[3]);
			result.Group = Convert.ToInt32(reader[4]);
			CloseConnection();
			return result;
		}

		public WorkingPosition GetWorkingPosition(int ID)
		{
			WorkingPosition result;
			//var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From WorkingPosition Where ID = @ID";
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new WorkingPosition(Convert.ToInt32(reader[0]));
			result.Name = reader[1].ToString();
			CloseConnection();
			return result;
		}

		public WorkloadAssign GetWorkloadAssign(int ID)
		{
			WorkloadAssign result;
			//var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From WorkloadAssign Where ID = @ID";
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new WorkloadAssign(Convert.ToInt32(reader[0]));
			result.EmployeeID = Convert.ToInt32(reader[1]);
			result.WorkloadID = Convert.ToInt32(reader[2]);
			result.StudentsCount = Convert.ToInt32(reader[3]);
			result.Weeks = Convert.ToInt32(reader[4]);
			CloseConnection();
			return result;
		}



		public StudyForm GetStudyForm(int ID)
		{
			StudyForm result;
			//var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From StudyForm Where ID = @ID";
			cmd.Parameters.AddWithValue("@ID", ID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			result = new StudyForm(Convert.ToInt32(reader[0]));
			result.Name = reader[1].ToString();
			result.NameRus = reader[2].ToString();
			CloseConnection();
			return result;
		}
		#endregion

		#region GetLists
		public List<AcademicDegree> GetAcademicDegrees()
		{
			List<AcademicDegree> result = new List<AcademicDegree>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From AcademicDegree";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				//DataRow query = DataSetHolder.GetInstance().SemestersQuery;
				//foreach (DataRow row in DataSetHolder.GetInstance().SemestersQuery)
				AcademicDegree item = new AcademicDegree(Convert.ToInt32(reader[0]));
				item.Name = reader[1].ToString();
				result.Add(item);
			}

			return result;
		}



		public List<AcademicRank> GetAcademicRanks()
		{
			List<AcademicRank> result = new List<AcademicRank>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From AcademicRank";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				AcademicRank item = new AcademicRank(Convert.ToInt32(reader[0]));
				item.Name = reader[1].ToString();
				result.Add(item);
			}

			return result;
		}

		public List<StudyYear> GetStudyYears()
		{
			List<StudyYear> result = new List<StudyYear>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From StudyYear";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				StudyYear year = new StudyYear(Convert.ToInt32(reader[0]));
				year.Year = Convert.ToInt32(reader[1]);
				result.Add(year);
			}

			return result;
		}

		public List<Department> GetDepartments()
		{
			List<Department> result = new List<Department>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Department";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				Department item = new Department(Convert.ToInt32(reader[0]));
				item.CodeNumber = Convert.ToInt32(reader[1]);
				item.Description = reader[1].ToString();
				result.Add(item);
			}

			return result;
		}

		public List<Discipline> GetDisciplines()
		{
			List<Discipline> result = new List<Discipline>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Discipline";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				Discipline item = new Discipline(Convert.ToInt32(reader[0]));
				item = new Discipline(Convert.ToInt32(reader[0]));
				item.Descr = reader[1].ToString();
				item.DepartmentID = Convert.ToInt32(reader[2]);
				item.TypeID = Convert.ToInt32(reader[3]);
				item.LectureCount = Convert.ToInt32(reader[4]);
				item.PracticeCount = Convert.ToInt32(reader[5]);
				item.LabCount = Convert.ToInt32(reader[6]);
				item.KR = Convert.ToBoolean(reader[7]);
				item.KP = Convert.ToBoolean(reader[8]);
				item.RGR = Convert.ToBoolean(reader[9]);
				item.Zach = Convert.ToBoolean(reader[10]);
				item.Ekz = Convert.ToBoolean(reader[11]);
				item.Kons = Convert.ToBoolean(reader[12]);
				item.UchPr = Convert.ToInt32(reader[13]);
				item.PrPr = Convert.ToInt32(reader[14]);
				item.PredDipPr = Convert.ToInt32(reader[15]);
				item.KonsZaoch = Convert.ToBoolean(reader[16]);
				item.GEK = Convert.ToBoolean(reader[17]);
				item.GAK = Convert.ToBoolean(reader[18]);
				item.GAKPred = Convert.ToBoolean(reader[19]);
				item.DPRuk = Convert.ToBoolean(reader[20]);
				item.DopuskVKR = Convert.ToBoolean(reader[21]);
				item.RetzVKR = Convert.ToBoolean(reader[22]);
				item.DPretz = Convert.ToBoolean(reader[23]);
				item.ASPRuk = Convert.ToBoolean(reader[24]);
				item.MAGRuk = Convert.ToBoolean(reader[25]);
				item.MAGRetz = Convert.ToBoolean(reader[26]);
				item.RukKaf = Convert.ToBoolean(reader[27]);
				item.NIIR = Convert.ToInt32(reader[28]);
				item.Special = Convert.ToBoolean(reader[29]);
				result.Add(item);
			}

			return result;
		}

		public List<Employee> GetEmployees()
		{
			List<Employee> result = new List<Employee>();
			//var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			OpenConnection();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Employee";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				Employee item = new Employee(Convert.ToInt32(reader[0]));
				item.Name = reader[1].ToString();
				if (reader[2] != System.DBNull.Value)
					item.ContractNumber = Convert.ToInt32(reader[2]);
				if (reader[3] != System.DBNull.Value)
					item.Rate = Convert.ToInt32(reader[3]);
				if (reader[4] != System.DBNull.Value)
					item.BirthDay = Convert.ToDateTime(reader[4]);
				if (reader[5] != System.DBNull.Value)
					item.WorkingPositionID = Convert.ToInt32(reader[5]);
				if (reader[6] != System.DBNull.Value)
					item.AcademicRankID = Convert.ToInt32(reader[6]);
				if (reader[7] != System.DBNull.Value)
					item.AcademicDegreeID = Convert.ToInt32(reader[7]);
				result.Add(item);
			}
			CloseConnection();
			return result;
		}

		public List<Faculty> GetFaculties()
		{
			List<Faculty> result = new List<Faculty>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Faculty";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				Faculty item = new Faculty(Convert.ToInt32(reader[0]));
				item.Name = reader[1].ToString();
				item.FullName = reader[2].ToString();
				result.Add(item);
			}

			return result;
		}

		public List<Group> GetGroups()
		{
			List<Group> result = new List<Group>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Group";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				Group item = new Group(Convert.ToInt32(reader[0]));
				item.Name = reader[1].ToString();
				item.Speciality = Convert.ToInt32(reader[2]);
				item.StudyQualification = Convert.ToInt32(reader[3]);
				item.StudyFormID = Convert.ToInt32(reader[4]);
				item.StudentCount = Convert.ToInt32(reader[5]);
				item.EntryYear = Convert.ToInt32(reader[6]);
				result.Add(item);
			}

			return result;
		}

		public List<Qualification> GetQualifications()
		{
			List<Qualification> result = new List<Qualification>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Qualification";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				Qualification item = new Qualification(Convert.ToInt32(reader[0]));
				item.Name = reader[1].ToString();
				result.Add(item);
			}

			return result;
		}

		public List<Semester> GetSemesters()
		{
			List<Semester> result = new List<Semester>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Semester";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				Semester item = new Semester(Convert.ToInt32(reader[0]));
				item.Name = reader[1].ToString();
				item.WeekCount = Convert.ToInt32(reader[2]);
				result.Add(item);
			}

			return result;
		}

		public List<Speciality> GetSpecialities()
		{
			List<Speciality> result = new List<Speciality>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Speciality";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				Speciality item = new Speciality(Convert.ToInt32(reader[0]));
				item.Name = reader[1].ToString();
				item.Faculty = Convert.ToInt32(reader[2]);
				result.Add(item);
			}

			return result;
		}

		public List<Student> GetStudents()
		{
			List<Student> result = new List<Student>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Student";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				Student item = new Student(Convert.ToInt32(reader[0]));
				item.Name = reader[1].ToString();
				result.Add(item);
			}

			return result;
		}

		public List<StudyForm> GetStudyForms()
		{
			List<StudyForm> result = new List<StudyForm>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From StudyForm";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				StudyForm item = new StudyForm(Convert.ToInt32(reader[0]));
				item.Name = reader[1].ToString();
				item.NameRus = reader[2].ToString();
				result.Add(item);
			}

			return result;
		}

		public List<WorkingPosition> GetWorkingPositions()
		{
			List<WorkingPosition> result = new List<WorkingPosition>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From WorkingPosition";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				WorkingPosition item = new WorkingPosition(Convert.ToInt32(reader[0]));
				item.Name = reader[1].ToString();
				result.Add(item);
			}

			return result;
		}

		public List<Workload> GetWorkloads()
		{
			List<Workload> result = new List<Workload>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From Workload";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				Workload item = new Workload(Convert.ToInt32(reader[0]));
				item.Discipline = Convert.ToInt32(reader[1]);
				item.Year = Convert.ToInt32(reader[2]);
				item.Semester = Convert.ToInt32(reader[3]);
				item.Group = Convert.ToInt32(reader[4]);
				result.Add(item);
			}

			return result;
		}


		public List<Workload> GetEmployeeWorkloads(int employeeID, int yearID)
		{
			List<Workload> result = new List<Workload>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			if (cn.State != ConnectionState.Open)
				cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "SELECT Workload.ID, Workload.Discipline, Workload.StudyYear, Workload.Semester, Workload.[Group] " +
								"FROM(WorkloadAssign INNER JOIN " +
								"Workload ON WorkloadAssign.Workload = Workload.ID) " +
								"WHERE(Workload.StudyYear = @yearID) AND(WorkloadAssign.Teacher = @employeeID)";

			cmd.Parameters.AddWithValue("@yearID", yearID);
			cmd.Parameters.AddWithValue("@employeeID", employeeID);



			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				Workload item = new Workload(Convert.ToInt32(reader[0]));
				item.Discipline = Convert.ToInt32(reader[1]);
				item.Year = Convert.ToInt32(reader[2]);
				item.Semester = Convert.ToInt32(reader[3]);
				item.Group = Convert.ToInt32(reader[4]);
				result.Add(item);
			}
			CloseConnection();
			return result;
		}

		public List<WorkloadAssign> GetWorkloadAssigns()
		{
			List<WorkloadAssign> result = new List<WorkloadAssign>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From WorkloadAssign";
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				WorkloadAssign item = new WorkloadAssign(Convert.ToInt32(reader[0]));
				item.EmployeeID = Convert.ToInt32(reader[1]);
				item.WorkloadID = Convert.ToInt32(reader[2]);
				item.StudentsCount = Convert.ToInt32(reader[3]);
				item.Weeks = Convert.ToInt32(reader[4]);
				result.Add(item);
			}

			cn.Close();
			return result;
		}

		public WorkloadAssign GetWorkloadAssign(int workloadID, string EmployeeName)
		{
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "SELECT WorkloadAssign.ID, WorkloadAssign.Teacher, WorkloadAssign.Workload, WorkloadAssign.StudentCount, WorkloadAssign.Weeks " +
								" FROM(WorkloadAssign INNER JOIN Employee ON WorkloadAssign.Teacher = Employee.ID) " +
								" WHERE(WorkloadAssign.Workload = @workID) AND(Employee.EmployeeName = @teachName)";
			cmd.Parameters.AddWithValue("@workID", workloadID);
			cmd.Parameters.AddWithValue("@teachID", EmployeeName);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			reader.Read();

			WorkloadAssign item = new WorkloadAssign(Convert.ToInt32(reader[0]));
			item.EmployeeID = Convert.ToInt32(reader[1]);
			item.WorkloadID = Convert.ToInt32(reader[2]);
			if (reader[3] != System.DBNull.Value)
				item.StudentsCount = Convert.ToInt32(reader[3]);
			if (reader[4] != System.DBNull.Value)
				item.Weeks = Convert.ToInt32(reader[4]);

			cn.Close();
			return item;
		}

		public WorkloadAssign GetWorkloadAssign(int workloadID, int EmployeeID)
		{
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "SELECT WorkloadAssign.ID, WorkloadAssign.Teacher, WorkloadAssign.Workload, WorkloadAssign.StudentCount, WorkloadAssign.Weeks " +
								" FROM(WorkloadAssign INNER JOIN Employee ON WorkloadAssign.Teacher = Employee.ID) " +
								" WHERE(WorkloadAssign.Workload = @workID) AND(Employee.ID = @EmployeeID)";
			cmd.Parameters.AddWithValue("@workID", workloadID);
			cmd.Parameters.AddWithValue("@EmployeeID", EmployeeID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			reader.Read();

			WorkloadAssign item = new WorkloadAssign(Convert.ToInt32(reader[0]));
			item.EmployeeID = Convert.ToInt32(reader[1]);
			item.WorkloadID = Convert.ToInt32(reader[2]);
			if (reader[3] != System.DBNull.Value)
				item.StudentsCount = Convert.ToInt32(reader[3]);
			if (reader[4] != System.DBNull.Value)
				item.Weeks = Convert.ToInt32(reader[4]);

			cn.Close();
			return item;
		}

		public List<WorkloadAssign> GetWorkloadAssigns(int workloadID)
		{
			List<WorkloadAssign> result = new List<WorkloadAssign>();
			var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			cn.Open();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From WorkloadAssign Where Workload = @ID";
			cmd.Parameters.AddWithValue("@ID", workloadID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			while (reader.Read())
			{
				WorkloadAssign item = new WorkloadAssign(Convert.ToInt32(reader[0]));
				item.EmployeeID = Convert.ToInt32(reader[1]);
				item.WorkloadID = Convert.ToInt32(reader[2]);
				if (reader[3] != System.DBNull.Value)
					item.StudentsCount = Convert.ToInt32(reader[3]);
				if (reader[4] != System.DBNull.Value)
					item.Weeks = Convert.ToInt32(reader[4]);
				result.Add(item);
			}
			cn.Close();
			return result;
		}

		public bool CheckWorkloadAssign(int workloadID)
		{
			bool res = false;

			//var cn = new System.Data.OleDb.OleDbConnection(connString);
			var cmd = new System.Data.OleDb.OleDbCommand();
			//cn.Open();
			OpenConnection();
			cmd.Connection = cn;
			cmd.CommandType = CommandType.Text;
			cmd.CommandText = "Select * From WorkloadAssign Where Workload = @ID";
			cmd.Parameters.AddWithValue("@ID", workloadID);
			System.Data.OleDb.OleDbDataReader reader = cmd.ExecuteReader();
			// while there is another record present
			res = reader.HasRows;
			CloseConnection();

			return res;
		}

		#endregion
		#endregion

		#region SetData
		public void AssignTeacher(int workloadID, int teacherID, int studCount = 0)
		{
			string insertSql = "INSERT INTO WorkloadAssign (Workload,Teacher, StudentCount) VALUES(@workloadID, @teacherID, @studentCount)";

			using (OleDbConnection myConnection = new OleDbConnection(connString))
			{
				myConnection.Open();

				OleDbCommand cmd = new OleDbCommand(insertSql, myConnection);

				cmd.Parameters.AddWithValue("@workloadID", workloadID);
				cmd.Parameters.AddWithValue("@teacherID", teacherID);
				cmd.Parameters.AddWithValue("@studentCount", studCount);
				cmd.ExecuteNonQuery();
				myConnection.Close();
			}
		}

		public void FreeTeacher(int workloadID, int teacherID)
		{
			string deleteSQL = "DELETE FROM WorkloadAssign WHERE(WorkloadAssign.Teacher = @teacherID) AND(WorkloadAssign.Workload = @workloadID)";

			using (OleDbConnection myConnection = new OleDbConnection(connString))
			{
				myConnection.Open();

				OleDbCommand cmd = new OleDbCommand(deleteSQL, myConnection);

				cmd.Parameters.AddWithValue("@teacherID", teacherID);
				cmd.Parameters.AddWithValue("@workloadID", workloadID);
				cmd.ExecuteNonQuery();
				myConnection.Close();
			}

		}
		public int AddDiscipline(Discipline discipline)
		{
			int id = -1;

			string insertSql =
	"INSERT INTO Discipline(Descr," +
	" Department," +
	" DisciplineType," +
	" LectureCount," +
	" PracticeCount," +
	" LabCount," +
	" KR," +
	" KP," +
	" RGR," +
	" Zach," +
	" Ekz," +
	" Kons," +
	" UchPr, " +
	" prPr," +
	" predDipPr," +
	" NIIR," +
	" KonsZaoch," +
	" GEK, " +
	" GAK," +
	" GosEkz," +
	" GAKPred," +
	" DPruk," +
	" DopuskVKR," +
	" RetzVKR," +
	" DPretz," +
	" ASPRuk," +
	" MAGRuk," +
	" MAGRetz," +
	" RukKaf," +
	" isSpecial," +
	" Contr) " +
	"VALUES(@Descr," +
	"@Department," +
	"@DisciplineType," +
	"@LectureCount," +
	"@PracticeCount," +
	"@LabCount," +
	"@KR," +
	"@KP," +
	"@RGR," +
	"@Zach," +
	"@Ekz," +
	"@Kons," +
	"@UchPr," +
	" @prPr," +
	" @predDipPr," +
	" @NIIR," +
	" @KonsZaoch," +
	" @GEK," +
	" @GAK," +
	" @GosEkz," +
	" @GAKPred," +
	" @DPruk," +
	" @DopuskVKR," +
	" @RetzVKR," +
	" @DPretz," +
	" @ASPRuk," +
	" @MAGRuk," +
	" @MAGRetz," +
	" @RukKaf," +
	" @isSpecial," +
	"@Contr)";

			using (OleDbConnection myConnection = new OleDbConnection(connString))
			{
				myConnection.Open();

				OleDbCommand cmd = new OleDbCommand(insertSql, myConnection);

				cmd.Parameters.AddWithValue("@Descr", discipline.Descr);
				cmd.Parameters.AddWithValue("@Department", discipline.DepartmentID);
				cmd.Parameters.AddWithValue("@DisciplineType", discipline.TypeID);
				cmd.Parameters.AddWithValue("@LectureCount", discipline.LectureCount);
				cmd.Parameters.AddWithValue("@PracticeCount", discipline.PracticeCount);
				cmd.Parameters.AddWithValue("@LabCount", discipline.LabCount);
				cmd.Parameters.AddWithValue("@KR", discipline.KR);
				cmd.Parameters.AddWithValue("@KP", discipline.KP);
				cmd.Parameters.AddWithValue("@RGR", discipline.RGR);
				cmd.Parameters.AddWithValue("@Zach", discipline.Zach);
				cmd.Parameters.AddWithValue("@Ekz", discipline.Ekz);
				cmd.Parameters.AddWithValue("@Kons", discipline.Kons);

				cmd.Parameters.AddWithValue("@UchPr", discipline.UchPr);
				cmd.Parameters.AddWithValue("@prPr", discipline.PrPr);
				cmd.Parameters.AddWithValue("@predDipPr", discipline.PredDipPr);
				cmd.Parameters.AddWithValue("@NIIR", discipline.NIIR);
				cmd.Parameters.AddWithValue("@KonsZaoch", discipline.KonsZaoch);
				cmd.Parameters.AddWithValue("@GEK", discipline.GEK);
				cmd.Parameters.AddWithValue("@GAK", discipline.GAK);
				cmd.Parameters.AddWithValue("@GosEkz", discipline.GosEkz);
				cmd.Parameters.AddWithValue("@GAKPred", discipline.GAKPred);
				cmd.Parameters.AddWithValue("@DPruk", discipline.DPRuk);
				cmd.Parameters.AddWithValue("@DopuskVKR", discipline.DopuskVKR);
				cmd.Parameters.AddWithValue("@RetzVKR", discipline.RetzVKR);
				cmd.Parameters.AddWithValue("@DPretz", discipline.DPretz);
				cmd.Parameters.AddWithValue("@ASPRuk", discipline.ASPRuk);
				cmd.Parameters.AddWithValue("@MAGRuk", discipline.MAGRuk);
				cmd.Parameters.AddWithValue("@MAGRetz", discipline.MAGRetz);
				cmd.Parameters.AddWithValue("@RukKaf", discipline.RukKaf);
				cmd.Parameters.AddWithValue("@isSpecial", discipline.Special);

                cmd.Parameters.AddWithValue("@Contr", discipline.Contract);
				cmd.ExecuteNonQuery();

				string query = "Select @@Identity";
				cmd.CommandText = query;
				id = (int)cmd.ExecuteScalar();
				myConnection.Close();
			}

			return id;
		}

		public int AddWorkload(Workload workload)
		{
			int id = -1;
			string insertSql =
	"INSERT INTO Workload (Discipline, StudyYear, Semester, [Group]) VALUES(@Discipline,@StudyYear,@Semester,@Group)";

			using (OleDbConnection myConnection = new OleDbConnection(connString))
			{
				myConnection.Open();

				OleDbCommand cmd = new OleDbCommand(insertSql, myConnection);

				cmd.Parameters.AddWithValue("@Discipline", workload.Discipline);
				cmd.Parameters.AddWithValue("@StudyYear", workload.Year);
				cmd.Parameters.AddWithValue("@Semester", workload.Semester);
				cmd.Parameters.AddWithValue("@Group", workload.Group);

				try
				{
					cmd.ExecuteNonQuery();
				}
				catch (Exception err)
				{
					System.Windows.Forms.MessageBox.Show(string.Format("@Discipline = {0}\n@StudyYear = {1}\n@Semester = {2}\n@Group = {3}\n", workload.Discipline, workload.Year, workload.Semester, workload.Group));
				}
				string query = "Select @@Identity";
				cmd.CommandText = query;
				id = (int)cmd.ExecuteScalar();
				myConnection.Close();
			}
			return id;
		}

		public int AddWorkloadAssign(WorkloadAssign assign)
		{
			int id = -1;
			string insertSql =
	"INSERT INTO WorkloadAssign (Teacher, Workload) VALUES(@Teacher,@Workload)";

			using (OleDbConnection myConnection = new OleDbConnection(connString))
			{
				myConnection.Open();

				OleDbCommand cmd = new OleDbCommand(insertSql, myConnection);

				cmd.Parameters.AddWithValue("@Discipline", assign.EmployeeID);
				cmd.Parameters.AddWithValue("@StudyYear", assign.WorkloadID);

				try
				{
					cmd.ExecuteNonQuery();
				}
				catch (Exception err)
				{
					//System.Windows.Forms.MessageBox.Show(string.Format("@Discipline = {0}\n@StudyYear = {1}\n@Semester = {2}\n@Group = {3}\n", workload.Discipline, workload.Year, workload.Semester, workload.Group));
				}
				string query = "Select @@Identity";
				cmd.CommandText = query;
				id = (int)cmd.ExecuteScalar();
				myConnection.Close();
			}
			return id;
		}

		public void SetAcademicDegree(AcademicDegree item)
		{
			using (OleDbConnection con = new OleDbConnection(connString))
			using (OleDbCommand command = con.CreateCommand())
			{
				command.CommandText = "UPDATE academicdegree SET Descr = ? WHERE ID = ?";
				command.Parameters.Add("@p1", OleDbType.VarChar).Value = item.Name;
				command.Parameters.Add("@p2", OleDbType.Integer).Value = item.ID;
				con.Open();
				command.ExecuteNonQuery();
			}
		}

		public void SetAcademicRank(AcademicRank item)
		{
			using (OleDbConnection con = new OleDbConnection(connString))
			using (OleDbCommand command = con.CreateCommand())
			{
				command.CommandText = "UPDATE AcademicRank SET Descr = ? WHERE ID = ?";
				command.Parameters.Add("@p1", OleDbType.VarChar).Value = item.Name;
				command.Parameters.Add("@p2", OleDbType.Integer).Value = item.ID;
				con.Open();
				command.ExecuteNonQuery();
			}
		}

		public void SetDepartment(Department item)
		{
			using (OleDbConnection con = new OleDbConnection(connString))
			using (OleDbCommand command = con.CreateCommand())
			{
				command.CommandText = "UPDATE Department SET Descr = ?, CodeNumber = ? WHERE ID = ?";
				command.Parameters.Add("@p1", OleDbType.VarChar).Value = item.Description;
				command.Parameters.Add("@p2", OleDbType.VarChar).Value = item.CodeNumber;
				command.Parameters.Add("@p3", OleDbType.Integer).Value = item.ID;
				con.Open();
				command.ExecuteNonQuery();
			}
		}
        
        public void SetSemester(Semester item)
        {
            using (OleDbConnection con = new OleDbConnection(connString))
            using (OleDbCommand command = con.CreateCommand())
            {
                command.CommandText = "UPDATE Semester Set WeekCount = ? Where ID = ?";
                command.Parameters.Add("@p1", OleDbType.Integer).Value = item.WeekCount;
                command.Parameters.Add("@p2", OleDbType.Integer).Value = item.ID;
                con.Open();
                command.ExecuteNonQuery();
            }
        }

		public void SetDiscipline(Discipline item)
		{
			using (OleDbConnection con = new OleDbConnection(connString))
			using (OleDbCommand command = con.CreateCommand())
			{
				command.CommandText = "UPDATE Discipline SET Descr = ?, Department = ?, DisciplineType = ? " +
				 "LectureCount=?, PracticeCount=?, LabCount=?, KR=?, KP=?, RGR=?,Zach=?,Ekz=?,Kons=? WHERE ID = ?";
				command.Parameters.Add("@p1", OleDbType.VarChar).Value = item.Descr;
				command.Parameters.Add("@p3", OleDbType.Integer).Value = item.DepartmentID;
				command.Parameters.Add("@p2", OleDbType.Integer).Value = item.TypeID;
				command.Parameters.Add("@p2", OleDbType.Integer).Value = item.LectureCount;
				command.Parameters.Add("@p2", OleDbType.Integer).Value = item.PracticeCount;
				command.Parameters.Add("@p2", OleDbType.Integer).Value = item.LabCount;
				command.Parameters.Add("@p2", OleDbType.Boolean).Value = item.KR;
				command.Parameters.Add("@p2", OleDbType.Boolean).Value = item.KP;
				command.Parameters.Add("@p2", OleDbType.Boolean).Value = item.RGR;
				command.Parameters.Add("@p2", OleDbType.Boolean).Value = item.Zach;
				command.Parameters.Add("@p2", OleDbType.Boolean).Value = item.Ekz;
				command.Parameters.Add("@p2", OleDbType.Boolean).Value = item.Kons;
				command.Parameters.Add("@p2", OleDbType.Integer).Value = item.ID;
				con.Open();
				command.ExecuteNonQuery();
			}
		}

		public void SetEmployee(Employee item)
		{
			using (OleDbConnection con = new OleDbConnection(connString))
			using (OleDbCommand command = con.CreateCommand())
			{
				command.CommandText = "UPDATE Employee " +
				"SET EmployeeName =?, Rate =?, ContractNumber =?, BirthDay =?, WorkingPosition =?, AcademicRank =?, AcademicDegree =? " +
				"WHERE(ID = ?)";
				command.Parameters.Add("@p1", OleDbType.VarChar).Value = item.Name;
				command.Parameters.Add("@p3", OleDbType.Single).Value = item.Rate;
				command.Parameters.Add("@p2", OleDbType.Integer).Value = item.ContractNumber;
				command.Parameters.Add("@p2", OleDbType.Date).Value = item.BirthDay;
				command.Parameters.Add("@p2", OleDbType.Integer).Value = item.WorkingPositionID;
				command.Parameters.Add("@p2", OleDbType.Integer).Value = item.AcademicRankID;
				command.Parameters.Add("@p2", OleDbType.Integer).Value = item.AcademicDegreeID;
				command.Parameters.Add("@p2", OleDbType.Integer).Value = item.ID;
				con.Open();
				command.ExecuteNonQuery();
			}
		}

		public void SetFaculty(Faculty item)
		{
			using (OleDbConnection con = new OleDbConnection(connString))
			using (OleDbCommand command = con.CreateCommand())
			{
				command.CommandText = "UPDATE Faculty " +
				"SET Descr =?, FullName =? " +
				"WHERE(ID = ?)";
				command.Parameters.Add("@p1", OleDbType.VarChar).Value = item.Name;
				command.Parameters.Add("@p3", OleDbType.Single).Value = item.FullName;
				con.Open();
				command.ExecuteNonQuery();
			}
		}

		public void SetGroup(Group item)
		{
			using (OleDbConnection con = new OleDbConnection(connString))
			using (OleDbCommand command = con.CreateCommand())
			{
				command.CommandText = "UPDATE Group " +
				"SET Descr = ?, Speciality = ?, Qualification = ?, StudyForm = ?, StudentCount = ?, EntryYear = ?" +
				"WHERE(ID = ?)";
				command.Parameters.Add("@p1", OleDbType.VarChar).Value = item.Name;
				command.Parameters.Add("@p2", OleDbType.Integer).Value = item.Speciality;
				command.Parameters.Add("@p3", OleDbType.Integer).Value = item.StudyQualification;
				command.Parameters.Add("@p4", OleDbType.Integer).Value = item.StudyFormID;
				command.Parameters.Add("@p5", OleDbType.Integer).Value = item.StudentCount;
				command.Parameters.Add("@p6", OleDbType.Integer).Value = item.EntryYear;
				command.Parameters.Add("@p7", OleDbType.Integer).Value = item.ID;
				con.Open();
				command.ExecuteNonQuery();
			}
		}

		#endregion

		public bool CheckDisciplinesAtYearExist(string year)
		{
			var cmd = CreateQuery("SELECT Workload.ID FROM(Workload INNER JOIN StudyYear ON Workload.StudyYear = StudyYear.ID) " +
									"WHERE(StudyYear.StudyYear = @YEAR) ");
			cmd.Parameters.AddWithValue("@YEAR", year);

			OleDbDataReader reader = cmd.ExecuteReader();

			reader.Read();

			bool result = reader.HasRows;

			CloseConnection();
			return result;
		}

	}

	public class DataSetHolder
	{
		public DataSetHolder()
		{
			mainDBDataset = new MainDBDataSet();
			mainDBDataset.DataSetName = "MainDBDataSet";
			mainDBDataset.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
			InitTablesFromDB();
		}

		#region singletone
		static DataSetHolder _instance;

		public static DataSetHolder GetInstance()
		{
			if (_instance == null)
				_instance = new DataSetHolder();

			return _instance;
		}

		#endregion
		MainDBDataSet mainDBDataset;

		public void InitTablesFromDB()
		{

		}

		public SemestersQueryDataTable SemestersQuery
		{
			get
			{
				((System.ComponentModel.ISupportInitialize)(mainDBDataset)).BeginInit();
				SemestersQueryTableAdapter semesterTableAdapter = new SemestersQueryTableAdapter();
				semesterTableAdapter.ClearBeforeFill = true;
				semesterTableAdapter.Fill(mainDBDataset.SemestersQuery);
				((System.ComponentModel.ISupportInitialize)(mainDBDataset)).EndInit();
				return mainDBDataset.SemestersQuery;

			}
		}

		public YearsQueryDataTable YearsQuery
		{
			get
			{
				YearsQueryTableAdapter yearsTableAdapter = new YearsQueryTableAdapter();
				yearsTableAdapter.ClearBeforeFill = true;
				yearsTableAdapter.Fill(mainDBDataset.YearsQuery);

				return mainDBDataset.YearsQuery;
			}
		}

		public YearsExistQueryDataTable YearsExistQuery
		{
			get
			{
				YearsExistQueryTableAdapter yearsTableAdapter = new YearsExistQueryTableAdapter();
				yearsTableAdapter.ClearBeforeFill = true;
				yearsTableAdapter.Fill(mainDBDataset.YearsExistQuery);

				return mainDBDataset.YearsExistQuery;
			}
		}
	}

}
