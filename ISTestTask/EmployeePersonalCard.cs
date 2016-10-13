using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ISTestTask
{
	public class EmployeePersonalCard
	{
		public static string FileName { get { return "employees.xml";} set { } }

		public string DateCreation { get; set; }

		public Employee Employee { get; set; }

		public int DocNumber { get; set; }

		public EmployeePersonalCard()
		{

		}

		public EmployeePersonalCard(Employee employee)
		{
			this.Employee = employee;
			this.DateCreation = DateTime.Now.ToString();
		}

		public string getName()
		{
			return Employee.FamilyName;
		}
	}
}
