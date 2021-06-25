using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DzhalevPavel_SofiaDraftingInterview.Controllers;
using Microsoft.Win32;
using ExcelDataReader;
using Xceed.Wpf.Toolkit;
using MessageBox = System.Windows.MessageBox;

namespace DzhalevPavel_SofiaDraftingInterview
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		private async void OnImport(object sender, RoutedEventArgs e)
		{

			if (UsersController.ChooseFile(out string fileName))
			{
				BusyIndicator.IsBusy = true;
				UsersGrid.ItemsSource = await Task.Run(() => UsersController.ImportUsers(fileName));
				BusyIndicator.IsBusy = false;
			}
			else
			{
				MessageBox.Show("An error occurred while choosing a file.");
			}
			
		}
	}
}
