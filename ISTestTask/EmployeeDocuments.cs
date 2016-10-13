using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ISTestTask
{
	public class EmployeeDocuments
	{
		public EmployeeDocuments()
		{

		}
		public EmployeeDocuments(string inn, string snils)
		{

			Inn = inn;
			Snils = snils;
		}
		public string Inn { get; set; }
		public string Snils { get; set; }
	}
}
