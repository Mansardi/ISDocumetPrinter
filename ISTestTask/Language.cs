using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ISTestTask
{
	public class Language
	{
		public string Name { get; set; }

		public string Level { get; set; }

		public Language()
		{

		}

		public Language(string name, string level)
		{
			Name = name;
			Level = level;
		}
	}
}
