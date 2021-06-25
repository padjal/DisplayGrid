using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using DzhalevPavel_SofiaDraftingInterview.Controllers;
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
				//Change task to test slow behaviour.
				//UsersGrid.ItemsSource = await Task.Run(() => UsersController.ImportUsersSlow(fileName));
				stopwatch.Stop();
				Timer.Text = $"Import completed in {stopwatch.Elapsed.TotalMilliseconds} milliseconds";
				BusyIndicator.IsBusy = false;
			}
			else MessageBox.Show("An error occurred while choosing a file.");
		}
	}
}
