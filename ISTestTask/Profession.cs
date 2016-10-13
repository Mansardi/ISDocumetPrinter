using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ISTestTask
{
	public class Profession
	{
		public Profession() { }

		public Profession(string main, string other)
		{
			Main = main;
			Other = other;
		}

		public string Main { get; set; }
		public string Other { get; set; }

	}
}
