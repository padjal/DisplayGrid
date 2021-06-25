using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DzhalevPavel_SofiaDraftingInterview.Controllers
{
	static class UsersController
	{
		public static bool ChooseFile(out string fileName)
		{
			var chooseFileDlog = new OpenFileDialog { Filter = "Excel files(*.xlsx;*.xls;*.xlt)|*.xlsx;*.xls;*.xlt" };

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

			// Open the document for editing.
			using (SpreadsheetDocument spreadsheetDocument =
				SpreadsheetDocument.Open(fileName, false))
			{
				WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
				WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
				SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

				foreach (Row r in sheetData.Elements<Row>())
				{
					if (r.RowIndex <= 2)
						continue;
					try
					{
						users.Add(ParseUser(spreadsheetDocument, r));
					}
					catch (Exception e)
					{
						MessageBox.Show($"Could not parse user. {e.Message}");
					}
				}
			}

			var sortedUsers = users.OrderBy(x => x.Name).ToList();

			return sortedUsers;
		}

		/// <summary>
		/// The first method which I used to complete the task. Left for diagnostics.
		/// </summary>
		/// <param name="fileName"></param>
		/// <returns></returns>
		public static List<User> ImportUsersSlow(string fileName)
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

				workBook.Close();
				excel.Quit();


			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message);
			}

			var sortedUsers = users.OrderBy(x => x.Name).ToList();

			return sortedUsers;
		}


		private static User ParseUser(SpreadsheetDocument document, Row row)
		{
			string[] user = new string[4];
			int index = 0;
			foreach (Cell c in row.Elements<Cell>())
			{
				user[index] = GetCellValue(document, c);
				index++;
			}
			return new User(user[0], user[1], user[2], user[3]);
		}

		private static string GetCellValue(SpreadsheetDocument document, Cell cell)
		{
			SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
			string value = cell.CellValue.InnerXml;

			if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
			{
				return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
			}
			else
			{
				return value;
			}
		}
	}

}
