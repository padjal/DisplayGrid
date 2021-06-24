using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDataReader;

namespace DzhalevPavel_SofiaDraftingInterview
{
	class User
	{
		public string Name { get; set; }
		public string Surname { get; set; }
		public string Location { get; set; }
		public string Email { get; set; }

		public User(string name, string surname, string location, string email)
		{
			Name = name;
			Surname = surname;
			Location = location;
			Email = email;
		}
	}

}
