using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ISTestTask
{
	public class Employee
	{
		public int TabNumber { get; set; }
		public string FamilyName { get; set; }
		public string Name { get; set; }
		public string SecondName { get; set; }
		public string DateBirth { get; set; }
		public string PlaceBirth { get; set; }
		public string Gender { get; set; }
		public Profession Profession { get; set; }
		public EmployeeDocuments Documets { get; set; }
		public string Work { get; set; }
		public string KindWork { get; set; }
		public string Nationality { get; set; }
		public Language Language { get; set; }
		public Education FirstEducation { get; set; }
		public Education SecondEducation { get; set; }
		public Education ScienceEducation { get; set; }
	}
}
