using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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

		/// <summary>
		/// Called when the import button is clicked. Initiates the importing of users.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private async void OnImport(object sender, RoutedEventArgs e)
		{

			if (UsersController.ChooseFile(out string fileName))
			{
				Stopwatch stopwatch = new Stopwatch();
				stopwatch.Start();

				BusyIndicator.IsBusy = true;
				UsersGrid.ItemsSource = await Task.Run(() => UsersController.ImportUsers(fileName));
				stopwatch.Stop();
				Timer.Text = $"Import completed in {stopwatch.Elapsed.TotalMilliseconds} milliseconds";
				BusyIndicator.IsBusy = false;
			}
			else MessageBox.Show("An error occurred while choosing a file.");
		}
	}
}
