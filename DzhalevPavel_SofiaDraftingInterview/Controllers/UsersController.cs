using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Windows;
using ExcelDataReader;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace DzhalevPavel_SofiaDraftingInterview.Controllers
{
	static class UsersController
	{
		public static bool ChooseFile(out string fileName)
		{
			var chooseFileDlog = new OpenFileDialog {Filter = "Excel files(*.xlsx;*.xls;*.xlt)|*.xlsx;*.xls;*.xlt"};

			if (chooseFileDlog.ShowDialog() is true)
			{
				fileName = chooseFileDlog.FileName;
				return true;
			}

			fileName = string.Empty;
			return false;
		}

		public static List<User> ImportUsers(string fileName)
		{
			List<User> users = new List<User>();

			Excel.Application excel = new Excel.Application();

			try
			{
				Excel.Workbook workBook = excel.Workbooks.Open(fileName);
				Excel.Worksheet excelSheet = excel.Worksheets[1];
				Excel.Range usedRange = excelSheet.UsedRange;

				for (int i = 3; i < usedRange.Rows.Count; i++)
				{
					string name = ((Excel.Range)usedRange.Cells[i, 1]).Value.ToString();
					string surname = ((Excel.Range)usedRange.Cells[i, 2]).Value2.ToString();
					string location = ((Excel.Range)usedRange.Cells[i, 3]).Value2.ToString();
					string email = ((Excel.Range)usedRange.Cells[i, 4]).Value2.ToString();
					users.Add(new User(name, surname, location, email));
				}
			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message);
			}

			var sortedUsers = users.OrderBy(x => x.Name).ToList();

			return sortedUsers;
		}
	}

}
