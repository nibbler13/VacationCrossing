using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace VacationCrossing {
	class ExcelHandlers {
		private ICellStyle cellStyleCrossing;
		private ICellStyle cellStyleNormal;
		private ISheet sheet;
		private IWorkbook workbook;
		private Dictionary<int, int> crossing = new Dictionary<int, int>();


		public static string AssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\";

		public static List<string> ReadSheetNames(string file) {
			List<string> sheetNames = new List<string>();

			using (OleDbConnection conn = new OleDbConnection()) {
				conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file + ";Mode=Read;" +
					"Extended Properties='Excel 12.0 Xml;HDR=NO;'";

				using (OleDbCommand comm = new OleDbCommand()) {
					conn.Open();
					DataTable dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
						new object[] { null, null, null, "TABLE" });
					foreach (DataRow row in dtSchema.Rows) {
						string name = row.Field<string>("TABLE_NAME");
						if (name.Contains("FilterDatabase"))
							continue;

						sheetNames.Add(name);
					}
				}
			}

			return sheetNames;
		}

		public static List<ItemEmployee> ReadExcelFile(string file, List<string> sheetNames, BackgroundWorker bw) {
			Dictionary<int, ItemEmployee> itemsInfo = new Dictionary<int, ItemEmployee>();
			bw.ReportProgress(0, "Считывание данных из файла: " + file);

			if (!File.Exists(file)) {
				bw.ReportProgress(0, "Файл не существует или нет доступа, пропуск обработки");
				return itemsInfo.Values.ToList();
			}

			try {
				using (OleDbConnection conn = new OleDbConnection()) {
					conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file + ";Mode=Read;" +
						"Extended Properties='Excel 12.0 Xml;HDR=NO;'";

					using (OleDbCommand comm = new OleDbCommand()) {
						conn.Open();
						DataTable dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
							new object[] { null, null, null, "TABLE" });
						List<string> sheetNamesInFile = new List<string>();
						foreach (DataRow row in dtSchema.Rows) {
							string sheetNameInFile = row.Field<string>("TABLE_NAME");
							bool isSelected = false;

							foreach (string sheetName in sheetNames) {
								if (sheetNameInFile.Contains(sheetName)) {
									isSelected = true;
									break;
								}
							}

							if (isSelected)
								sheetNamesInFile.Add(sheetNameInFile);
						}

						bw.ReportProgress(0, "Файл содержит листов: " + sheetNamesInFile.Count);

						foreach (string sheetName in sheetNamesInFile) {
							bw.ReportProgress(0, "Чтение содержимого листа: " + sheetName);
#pragma warning disable CA2100 // Review SQL queries for security vulnerabilities
							comm.CommandText = "Select * from [" + sheetName + "]";
#pragma warning restore CA2100 // Review SQL queries for security vulnerabilities
							comm.Connection = conn;

							DataTable dataTable = new DataTable();
							using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter()) {
								oleDbDataAdapter.SelectCommand = comm;
								oleDbDataAdapter.Fill(dataTable);
							}

							try {
								List<ItemEmployee> items = ParseDataTable(sheetName, dataTable, bw);
								if (items == null)
									bw.ReportProgress(0, "Не удалось считать данные с листа: " + sheetName);
								else
									foreach (ItemEmployee item in items) 
										if (!itemsInfo.ContainsKey(item.ID))
											itemsInfo.Add(item.ID, item);
							} catch (Exception e) {
								bw.ReportProgress(0, "Лист: " + sheetName + ", не удалось считать данные: " + e.Message);
							}
						}

						conn.Close();
					}
				}
			} catch (Exception e) {
				bw.ReportProgress(0, e.Message + Environment.NewLine + e.StackTrace);
			}

			return itemsInfo.Values.ToList();
		}


		private static List<ItemEmployee> ParseDataTable(string sheetName, DataTable dataTable, BackgroundWorker bw) {
			Dictionary<int, ItemEmployee> items = new Dictionary<int, ItemEmployee>();

			if (dataTable == null)
				return null;

			int columnsNeeded = 7;
			if (dataTable.Columns.Count < columnsNeeded) {
				bw.ReportProgress(0, "Кол-во столбцов меньше " + columnsNeeded + ", пропуск обработки");
				return null;
			}

			for (int i = 0; i < dataTable.Rows.Count; i++) {
				DataRow dataRow = dataTable.Rows[i];

				if (i == 0) {
					string searchColumn0 = dataRow[i].ToString().ToLower();
					if (!searchColumn0.StartsWith("подразделение организации")) {
						bw.ReportProgress(0, "Лист не соответствует формату, заголовок столбца A должен содержать 'Подразделение организации', пропуск обработки");
						return items.Values.ToList();
					}
				}

				string id = dataRow[3].ToString();
				if (string.IsNullOrEmpty(id)) {
					bw.ReportProgress(0, "Строка: " + (i + 1) + ", Таб. № - пусто, пропуск");
					continue;
				}

				int idValue = -1;
				if (!int.TryParse(id, out idValue)) {
					bw.ReportProgress(0, "Строка: " + (i + 1) + ", значение Таб. № не является числом, пропуск");
					continue;
				}

				if (!items.ContainsKey(idValue)) {
					ItemEmployee itemEmployee = new ItemEmployee {
						Department = dataRow[0].ToString(),
						Name = dataRow[1].ToString(),
						Position = dataRow[2].ToString(),
						ID = idValue,
						Type = dataRow[4].ToString()
					};

					items.Add(idValue, itemEmployee);
				}

				string vacationDays = dataRow[5].ToString();
				string vacationDateStart = dataRow[6].ToString().TrimStart(' ' ).TrimEnd(' ');
				if (string.IsNullOrEmpty(vacationDays) &&
					string.IsNullOrEmpty(vacationDateStart))
					continue;

				bool isVacationDaysParsed = int.TryParse(vacationDays, out int vacationDaysValue);
				bool isVacationDateStartParsed = DateTime.TryParse(vacationDateStart, out DateTime vacationDateStartValue);

				if (!isVacationDateStartParsed) {
					string[] splitted = null;

					if (vacationDateStart.Contains("-")) 
						splitted = vacationDateStart.Split('-');
					else if (vacationDateStart.Contains(" ")) 
						splitted = vacationDateStart.Split(' ');

					if (splitted != null) 
						isVacationDateStartParsed = DateTime.TryParse(
							splitted[0].TrimStart('.').TrimEnd('.'), out vacationDateStartValue);
				}

				if (!isVacationDaysParsed ||
					!isVacationDateStartParsed) {
					bw.ReportProgress(0, "!!! Строка: " + (i + 1) + ", не удалось разобрать кол-во календарных дней (" + 
						vacationDays + ") или дата запланированная (" + vacationDateStart +"), пропуск");
					continue;
				}

				items[idValue].vacationPeriods.Add(new Tuple<int, DateTime>(vacationDaysValue, vacationDateStartValue));
			}

			return items.Values.ToList();
		}




		public string WriteItemsToExcel(List<ItemEmployee> items, BackgroundWorker bw, int year) {
			workbook = null;
			sheet = null;
			string resultFile = string.Empty;
			string sheetName = "Данные";
			string resultFilePrefix = "Пересечение отпусков";
			string templateFileName = "Template.xlsx";

			foreach (ItemEmployee item in items) {
				foreach (Tuple<int, DateTime> vacationPeriod in item.vacationPeriods) {
					for (int i = 0; i < vacationPeriod.Item1; i++) {
						DateTime dateTime = vacationPeriod.Item2.AddDays(i);
						int dayOfYear = dateTime.DayOfYear;
						if (!crossing.ContainsKey(dayOfYear))
							crossing.Add(dayOfYear, 0);

						crossing[dayOfYear]++;
					}
				}
			}

			bw.ReportProgress(0, "Запись результатов в файл");
			bw.ReportProgress(0, "Создание новой книги Excel");
			try {
				if (!CreateNewIWorkbook(resultFilePrefix, templateFileName,
					out workbook, out sheet, out resultFile, sheetName))
					return string.Empty;
			} catch (Exception e) {
				bw.ReportProgress(0, e.Message);
				return string.Empty;
			}

			int columnNumber = 7;
			IFont font = workbook.CreateFont();
			font.FontHeightInPoints = 8;
			font.FontName = "Verdana";

			cellStyleCrossing = workbook.CreateCellStyle();
			cellStyleCrossing.SetFont(font);
			cellStyleCrossing.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightOrange.Index;
			cellStyleCrossing.FillPattern = FillPattern.SolidForeground;

			cellStyleNormal = workbook.CreateCellStyle();
			cellStyleNormal.SetFont(font);
			cellStyleNormal.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
			cellStyleNormal.FillPattern = FillPattern.SolidForeground;

			IRow rowMonth = GetRow(0);
			for (int month = 1; month <= 12; month++) 
				for (int day = 1; day <= DateTime.DaysInMonth(year, month); day++) {
					WriteValueToCell(month, columnNumber, rowMonth);
					columnNumber++;
				}			

			int rowNumber = 1;
			bw.ReportProgress(0, "Строк для записи: " + items.Count);
			foreach (ItemEmployee item in items) {
				double totalDays = 0;

				foreach (Tuple<int, DateTime> vacationPeriod in item.vacationPeriods) 
					totalDays += vacationPeriod.Item1;

				object[] emplyeeInfo = new object[] { 
					item.Department,
					item.Name,
					item.Position,
					item.ID,
					item.Type,
					totalDays
				};
				WriteValuesToRow(ref rowNumber, emplyeeInfo, item);
			}

			bw.ReportProgress(0, "Сохранение книги Excel");
			try {
				if (!SaveAndCloseIWorkbook(workbook, resultFile))
					return string.Empty;
			} catch (Exception e) {
				bw.ReportProgress(0, e.Message);
				return string.Empty;
			}

			return resultFile;
		}

		private IRow GetRow(int rowNumber) {
			IRow row = null;
			try { row = sheet.GetRow(rowNumber); } catch (Exception) { }

			if (row == null)
				row = sheet.CreateRow(rowNumber);

			return row;
		}

		private void WriteValuesToRow(ref int rowNumber, object[] values, ItemEmployee item) {
			IRow row = GetRow(rowNumber++);

			int column = 0;
			foreach (object value in values) {
				WriteValueToCell(value, column, row);
				column++;
			}

			foreach (Tuple<int, DateTime> vacationPeriod in item.vacationPeriods) 
				for (int i = 0; i < vacationPeriod.Item1; i++) {
					int dayOfYear = vacationPeriod.Item2.AddDays(i).DayOfYear;
					ICellStyle cellStyleToApply = cellStyleNormal;
					if (crossing[dayOfYear] > 1) {
						cellStyleToApply = cellStyleCrossing;
						WriteValueToCell("Да", 6, row);
					}

					WriteValueToCell(1, 6 + dayOfYear, row,
						vacationPeriod.Item2.ToShortDateString() + "(" + vacationPeriod.Item1 + ")", rowNumber - 1, cellStyleToApply);
				}
			
		}

		private void WriteValueToCell(object value, int columnNumber, IRow row, 
			string commentStr = null, int rowNumber = -1, ICellStyle cellStyle = null) {
			NPOI.SS.UserModel.ICell cell = null;
			try { cell = row.GetCell(columnNumber); } catch (Exception) { }

			if (cell == null)
				cell = row.CreateCell(columnNumber);

			if (double.TryParse(value.ToString(), out double valueDouble)) {
				cell.SetCellValue(valueDouble);
			} else if (value is DateTime) {
				cell.SetCellValue((DateTime)value);
			} else if (value == null) {
				return;
			} else {
				cell.SetCellValue(value.ToString());
			}

			if (!string.IsNullOrEmpty(commentStr) && sheet != null) {
				IDrawing patriarch = (XSSFDrawing)sheet.CreateDrawingPatriarch();
				IComment comment = patriarch.CreateCellComment(new XSSFClientAnchor());
				comment.Author = "VacationCrossing";
				comment.String = new XSSFRichTextString($"{comment.Author}:{Environment.NewLine}" + commentStr);
				comment.Visible = false;
				comment.Row = rowNumber;
				comment.Column = columnNumber;
				cell.CellComment = comment;
				cell.CellStyle = cellStyle;
			}
		}

		private static bool CreateNewIWorkbook(string resultFilePrefix, string templateFileName,
			out IWorkbook workbook, out ISheet sheet, out string resultFile, string sheetName) {
			workbook = null;
			sheet = null;
			resultFile = string.Empty;

			try {
				if (!GetTemplateFilePath(ref templateFileName))
					return false;

				resultFile = GetResultFilePath(resultFilePrefix, templateFileName);

				using (FileStream stream = new FileStream(templateFileName, FileMode.Open, FileAccess.Read))
					workbook = new XSSFWorkbook(stream);

				if (string.IsNullOrEmpty(sheetName))
					sheetName = "Данные";

				sheet = workbook.GetSheet(sheetName);

				return true;
			} catch (Exception) {
				throw;
			}
		}

		private static bool GetTemplateFilePath(ref string templateFileName) {
			templateFileName = Path.Combine(AssemblyDirectory, templateFileName);

			if (!File.Exists(templateFileName))
				return false;

			return true;
		}

		public static string GetResultFilePath(string resultFilePrefix, string templateFileName = "", bool isPlainText = false) {
			string resultPath = Path.Combine(AssemblyDirectory, "Results");
			if (!Directory.Exists(resultPath))
				Directory.CreateDirectory(resultPath);

			foreach (char item in Path.GetInvalidFileNameChars())
				resultFilePrefix = resultFilePrefix.Replace(item, '-');

			string fileEnding = ".xlsx";
			if (isPlainText)
				fileEnding = ".txt";

			string resultFile = Path.Combine(resultPath, resultFilePrefix + " " + DateTime.Now.ToString("yyyyMMdd_HHmmss") + fileEnding);

			if (isPlainText && !string.IsNullOrEmpty(templateFileName))
				File.Copy(templateFileName, resultFile, true);

			return resultFile;
		}

		private static bool SaveAndCloseIWorkbook(IWorkbook workbook, string resultFile) {
			try {
				using (FileStream stream = new FileStream(resultFile, FileMode.Create, FileAccess.Write))
					workbook.Write(stream);

				workbook.Close();

				return true;
			} catch (Exception) {
				throw;
			}
		}





		//============================ Interop Excel ============================
		private static bool OpenWorkbook(string workbook, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws, string sheetName = "") {
			xlApp = null;
			wb = null;
			ws = null;

			xlApp = new Excel.Application();

			if (xlApp == null) 
				throw new Exception("Не удалось открыть приложение Excel");

			xlApp.Visible = false;

			wb = xlApp.Workbooks.Open(workbook);

			if (wb == null)
				throw new Exception("Не удалось открыть книгу " + workbook);

			if (string.IsNullOrEmpty(sheetName))
				sheetName = "Данные";

			ws = wb.Sheets[sheetName];

			if (ws == null)
				throw new Exception("Не удалось открыть лист Данные");

			return true;
		}

		private static void SaveAndCloseWorkbook(Excel.Application xlApp, Excel.Workbook wb, Excel.Worksheet ws) {
			if (ws != null) {
				Marshal.ReleaseComObject(ws);
				ws = null;
			}

			if (wb != null) {
				wb.Save();
				wb.Close(0);
				Marshal.ReleaseComObject(wb);
				wb = null;
			}

			if (xlApp != null) {
				xlApp.Quit();
				Marshal.ReleaseComObject(xlApp);
				xlApp = null;
			}

			GC.Collect();
			GC.WaitForPendingFinalizers();
		}

		public static bool CopyFormatting(string resultFile) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws))
				return false;

			int rowsUsed = ws.UsedRange.Rows.Count;
			string lastColumn = GetExcelColumnName(ws.UsedRange.Columns.Count);

			ws.Range["A2:" + lastColumn + "2"].Select();
			xlApp.Selection.Copy();
			ws.Range["A3:" + lastColumn + rowsUsed].Select();
			xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
			ws.Rows["2:" + rowsUsed].Select();
			xlApp.Selection.RowHeight = 15;

			ws.Range["A1"].Select();

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static string GetExcelColumnName(int columnNumber) {
			int dividend = columnNumber;
			string columnName = String.Empty;
			int modulo;

			while (dividend > 0) {
				modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (int)((dividend - modulo) / 26);
			}

			return columnName;
		}


		private static void AddBoldBorder(Excel.Range range) {
			foreach (Excel.XlBordersIndex item in new Excel.XlBordersIndex[] {
					Excel.XlBordersIndex.xlEdgeBottom,
					Excel.XlBordersIndex.xlEdgeLeft,
					Excel.XlBordersIndex.xlEdgeRight,
					Excel.XlBordersIndex.xlEdgeTop}) {
				range.Borders[item].LineStyle = Excel.XlLineStyle.xlContinuous;
				range.Borders[item].ColorIndex = 0;
				range.Borders[item].TintAndShade = 0;
				range.Borders[item].Weight = Excel.XlBorderWeight.xlMedium;
			}
		}

		private static void AddInteriorColor(Excel.Range range, Excel.XlThemeColor xlThemeColor, double tintAndShade = 0.799981688894314) {
			range.Interior.Pattern = Excel.Constants.xlSolid;
			range.Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
			range.Interior.ThemeColor = xlThemeColor;
			range.Interior.TintAndShade = tintAndShade;
			range.Interior.PatternTintAndShade = 0;
		}

		public static string ColumnIndexToColumnLetter(int colIndex) {
			int div = colIndex;
			string colLetter = string.Empty;
			int mod = 0;

			while (div > 0) {
				mod = (div - 1) % 26;
				colLetter = (char)(65 + mod) + colLetter;
				div = (int)((div - mod) / 26);
			}

			return colLetter;
		}

		public static bool Process(string resultFile) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws))
				return false;

			xlApp.DisplayAlerts = false;

			int columnUsed = ws.UsedRange.Columns.Count;
			int rowUsed = ws.UsedRange.Rows.Count;

			string value = string.Empty;
			int rangeStart = 0;
			for (int column = 8; column <= columnUsed + 1; column++) {
				if (string.IsNullOrEmpty(value)) {
					object valueObj = ws.Cells[1, column].Value2;
					if (valueObj != null)
						value = valueObj.ToString();
					else
						value = string.Empty;

					rangeStart = column;
				}

				string valueNext = string.Empty;
				object valueNextObj = ws.Cells[1, column].value2;
				if (valueNextObj != null)
					valueNext = valueNextObj.ToString();

				if (!valueNext.Equals(value)) {
					ws.Cells[1, rangeStart].Value2 = GetMonthName(value);
					string rangeToMerge = ColumnIndexToColumnLetter(rangeStart) + "1:" + ColumnIndexToColumnLetter(column - 1) + "1";
					string rangeToAddBorder = rangeToMerge.Substring(0, rangeToMerge.Length - 1) + rowUsed;
					ws.Range[rangeToMerge].Merge();
					ws.Range[rangeToMerge].HorizontalAlignment = Excel.Constants.xlCenter;
					AddBoldBorder(ws.Range[rangeToAddBorder]);
					rangeStart = column;
					value = valueNext;
				}
			}

			ws.Columns["A:" + ColumnIndexToColumnLetter(columnUsed)].Select();
			xlApp.Selection.Font.Size = 8;
			ws.Columns["H:" + ColumnIndexToColumnLetter(columnUsed)].Select();
			xlApp.Selection.ColumnWidth = 0.5;

			ws.Range["H2"].Select();
			xlApp.ActiveWindow.FreezePanes = true;
			xlApp.Selection.AutoFilter();

			ws.Range["A1"].Select();

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static string GetMonthName(string month) {
			switch (month) {
				case "1":
					return "Январь";
				case "2":
					return "Февраль";
				case "3":
					return "Март";
				case "4":
					return "Апрель";
				case "5":
					return "Май";
				case "6":
					return "Июнь";
				case "7":
					return "Июль";
				case "8":
					return "Август";
				case "9":
					return "Сентябрь";
				case "10":
					return "Октябрь";
				case "11":
					return "Ноябрь";
				case "12":
					return "Декабрь";
				default:
					return "Ошибка";
			}
		}
	}
}
