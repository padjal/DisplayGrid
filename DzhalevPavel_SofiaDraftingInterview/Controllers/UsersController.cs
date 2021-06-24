using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace DzhalevPavel_SofiaDraftingInterview.Controllers
{
	class UsersController
	{

		public static List<User> ImportUsers()
		{
			List<User> users = new List<User>();

			var chooseFileDlog = new OpenFileDialog();
			chooseFileDlog.Filter = "Excel files(*.xlsx;*.xls;*.xlt)|*.xlsx;*.xls;*.xlt";
			if (chooseFileDlog.ShowDialog() is true)
			{
				string fileName = chooseFileDlog.FileName;

				Excel.Application excel = new Excel.Application();
				var workBook = excel.Workbooks.Open(fileName);
				Excel.Worksheet excelSheet = excel.Worksheets[1];
				Excel.Range usedRange = excelSheet.UsedRange;

				for (int i = 3; i < usedRange.Rows.Count; i++)
				{
					string name = ((Excel.Range)usedRange.Cells[i, 1]).Value2.ToString();
					string surname = ((Excel.Range)usedRange.Cells[i, 2]).Value2.ToString();
					string location = ((Excel.Range)usedRange.Cells[i, 3]).Value2.ToString();
					string email = ((Excel.Range)usedRange.Cells[i, 4]).Value2.ToString();
					users.Add(new User(name, surname, location, email));
				}
				
			}

			return users;
		}
	}

}
