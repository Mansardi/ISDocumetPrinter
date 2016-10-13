using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ISTestTask
{
	public class Education
	{
		
		public Education()
		{
			UniversityName = "";
			AcademicDegree = "";
			DiplomaName = "";
			DiplomaSeries = "";
			DiplomaNumber = "";
			Specialty = "";
			YearOfEnd = "";
		}

		public string UniversityName { get; set; }
		public string AcademicDegree { get; set; }
		public string DiplomaName { get; set; }
		public string DiplomaSeries { get; set; }
		public string DiplomaNumber { get; set; }
		public string Specialty { get; set; }
		public string YearOfEnd { get; set; }
	}
}
