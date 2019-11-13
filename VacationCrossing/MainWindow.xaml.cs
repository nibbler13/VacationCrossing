using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
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

namespace VacationCrossing {
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window {
		public int Year { get; set; }
		public string SelectedFile { get; set; }

		public ObservableCollection<string> SheetsAvailable { get; set; } = new ObservableCollection<string>();
		public ObservableCollection<string> SheetsSelected { get; set; } = new ObservableCollection<string>();

		public MainWindow() {
			InitializeComponent();

			Year = DateTime.Now.Year + 1;
			DataContext = this;
		}

		private void ButtonSelectFile_Click(object sender, RoutedEventArgs e) {
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.CheckFileExists = true;
			openFileDialog.CheckPathExists = true;
			openFileDialog.Filter = "Книга Excel (*.xls*)|*.xls*";
			openFileDialog.Multiselect = false;
			openFileDialog.Title = "Выберите книгу Excel, содержащие технокарты";

			if (openFileDialog.ShowDialog() == true) {
				TextBoxSelectedFile.Text = openFileDialog.FileName;
				SheetsAvailable.Clear();
				SheetsSelected.Clear();

				try {
					ExcelHandlers.ReadSheetNames(SelectedFile).ForEach(SheetsAvailable.Add);
					ButtonOneToSelected.IsEnabled = false;
					ButtonAllToSelected.IsEnabled = SheetsAvailable.Count > 0;
					ButtonAllToAvailable.IsEnabled = false;
					ButtonOneToAvailable.IsEnabled = false;
				} catch (Exception exc) {
					MessageBox.Show(this, exc.Message + Environment.NewLine + exc.StackTrace, "", MessageBoxButton.OK, MessageBoxImage.Error);
				}
			}
		}

		private void ButtonCreate_Click(object sender, RoutedEventArgs e) {
			string error = string.Empty;

			if (string.IsNullOrEmpty(SelectedFile))
				error = "Не выбран файл с графиком отпусков";

			if (SheetsSelected.Count == 0)
				error = "Не выбраны листы для формирования";

			string year = TextBoxYear.Text;
			if (string.IsNullOrEmpty(year) || !int.TryParse(year, out int yearValue))
				error = "Не указан год";

			if (!string.IsNullOrEmpty(error)) {
				MessageBox.Show(this, error, "", MessageBoxButton.OK, MessageBoxImage.Information);
				return;
			}

			BackgroundWorker bw = new BackgroundWorker();
			bw.DoWork += Bw_DoWork;
			bw.WorkerReportsProgress = true;
			bw.ProgressChanged += Bw_ProgressChanged;
			bw.RunWorkerCompleted += Bw_RunWorkerCompleted;
			bw.RunWorkerAsync();

			GridMain.Visibility = Visibility.Hidden;
			GridResults.Visibility = Visibility.Visible;
			TextBoxResult.Text = string.Empty;
			ButtonClose.IsEnabled = false;
			Cursor = Cursors.Wait;
		}

		private void Bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
			if (e.Error != null) {
				MessageBox.Show(this, "Во время выполнения произошла ошибка:" +
					Environment.NewLine + e.Error.Message + Environment.NewLine + e.Error.StackTrace, "", MessageBoxButton.OK, MessageBoxImage.Error);
			} else {
				MessageBox.Show(this, "Выполнение завершено", "",
					MessageBoxButton.OK, MessageBoxImage.Information);
			}

			ButtonClose.IsEnabled = true;
			Cursor = Cursors.Arrow;
		}

		private void Bw_ProgressChanged(object sender, ProgressChangedEventArgs e) {
			if (e.UserState is string) {
				TextBoxResult.Text += DateTime.Now.ToLongTimeString() + ": " + e.UserState + Environment.NewLine;
				TextBoxResult.ScrollToEnd();
			}
		}

		private void Bw_DoWork(object sender, DoWorkEventArgs e) {
			BackgroundWorker bw = sender as BackgroundWorker;
			List<ItemEmployee> itemEmployees = ExcelHandlers.ReadExcelFile(SelectedFile, SheetsSelected.ToList(), bw);
			bw.ReportProgress(0, "Сотрудников считано: " + itemEmployees.Count);

			if (itemEmployees.Count > 0) {
				ExcelHandlers excelHandlers = new ExcelHandlers();
				string resultFile = excelHandlers.WriteItemsToExcel(itemEmployees, bw, Year);
				bw.ReportProgress(0, "Применение форматирования (может занимать несколько минут)");
				if (ExcelHandlers.Process(resultFile))
					Process.Start(resultFile);
			}
		}

		private void ButtonClose_Click(object sender, RoutedEventArgs e) {
			GridResults.Visibility = Visibility.Hidden;
			GridMain.Visibility = Visibility.Visible;
		}

		private void ListView_SelectionChanged(object sender, SelectionChangedEventArgs e) {
			if (sender == ListViewAvailable) {
				ButtonOneToSelected.IsEnabled = ListViewAvailable.SelectedItems.Count > 0;
			} else {
				ButtonOneToAvailable.IsEnabled = ListViewSelected.SelectedItems.Count > 0;
			}
		}

		private void ListView_MouseDoubleClick(object sender, MouseButtonEventArgs e) {
			if (sender == ListViewAvailable)
				ButtonSheetNames_Click(ButtonOneToSelected, null);
			else
				ButtonSheetNames_Click(ButtonOneToAvailable, null);
		}

		private void ButtonSheetNames_Click(object sender, RoutedEventArgs e) {
			if (sender == ButtonOneToSelected) {
				string selected = ListViewAvailable.SelectedItem as string;
				SheetsAvailable.Remove(selected);
				SheetsSelected.Add(selected);
			} else if (sender == ButtonAllToSelected) {
				SheetsAvailable.ToList().ForEach(SheetsSelected.Add);
				SheetsAvailable.Clear();
			} else if (sender == ButtonAllToAvailable) {
				SheetsSelected.ToList().ForEach(SheetsAvailable.Add);
				SheetsSelected.Clear();
			} else if (sender == ButtonOneToAvailable) {
				string selected = ListViewSelected.SelectedItem as string;
				SheetsSelected.Remove(selected);
				SheetsAvailable.Add(selected);
			}

			ButtonAllToSelected.IsEnabled = SheetsAvailable.Count > 0;
			ButtonAllToAvailable.IsEnabled = SheetsSelected.Count > 0;
		}
	}
}
