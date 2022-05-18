namespace CafedralDB.SourceCode.Model
{
	/// <summary>
	/// Тип дисциплины
	/// </summary>
	enum DisciplineTypeEnum
	{

		Simple, //Обычная дисциплина 
		Special //Особая дисциплина с расчетом на человека(ГЭК,ГАК, Руководство и т.д.)
	}

	public class DisciplineType
	{
        public DisciplineType(int id)
		{
			ID = id;
		}

        public string Name { get; set; }

        public int ID { get; set; }
    }
}