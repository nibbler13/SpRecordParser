using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;

namespace SpRecordParser {
	class ExcelWriter {
		private ProgressBar progressBar;
		private System.Windows.Forms.TextBox textBox;
		Dictionary<string, SpRecordFileInformation> filesInfo;
		int percentage;

		public ExcelWriter(
			ProgressBar progressBar, 
			System.Windows.Forms.TextBox textBox,
			Dictionary<string, SpRecordFileInformation> filesInfo) {
			this.progressBar = progressBar;
			this.textBox = textBox;
			this.filesInfo = filesInfo;
			percentage = 20;
		}

		public bool WriteToExcel() {
			try {
				UpdateTextBox("Запуск приложения MS Excel");
				Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

				if (xlApp == null) {
					LoggingSystem.LogMessageToFile("Не удалось запустить MS Excel");
					return false;
				}

				xlApp.Visible = false;
				xlApp.EnableAnimations = false;
				xlApp.DisplayAlerts = false;

				UpdateTextBox("Создание новой книги");
				Workbook wb = xlApp.Workbooks.Add();
				foreach(Worksheet sheet in wb.Sheets) {
					if (sheet.Index != 1)
						sheet.Delete();
				}

				Worksheet ws = (Worksheet)wb.ActiveSheet;

				if (ws == null) {
					xlApp.Quit();
					LoggingSystem.LogMessageToFile("Не удалось создать Excel книгу");
					return false;
				}

				int oneFilePercentage = 70 / filesInfo.Count;
				foreach(KeyValuePair<string, SpRecordFileInformation> fileInfo in filesInfo) {
					List<List<string>> fileContent = fileInfo.Value.fileContent;
					UpdateTextBox("Заполнение информации по файлу: " + fileInfo.Key);

					int rowToFill = 1;
					foreach(List<string> line in fileContent) {
						string columnName = GetExcelColumnName(line.Count);
						ws.Range["A" + rowToFill + ":" + columnName + rowToFill].Value = line.ToArray();
						
						//line lenght = 9 - normal
						//			  = 10 - missed
						//			  = 11 - tied with missed

						if (line.Count > 9) {
							int colorIndex = 6;
							if (line.Count == 11)
								colorIndex = 35;
							ws.Range["A" + rowToFill + ":" + columnName + rowToFill].Interior.ColorIndex = colorIndex;
						}
						
						rowToFill++;
					}

					ws.UsedRange.Columns.AutoFit();
					ws.Columns[1].ColumnWidth = 15;

					wb.Worksheets.Add(After: ws);
					ws = (Worksheet)wb.ActiveSheet;
					percentage += oneFilePercentage;
					UpdateProgressBar(percentage);
				}

				UpdateTextBox("Создание сводной таблицы по всем файлам");
				ws.Range["A5"].Value = "Всего звонков";
				ws.Range["A5:A6"].Merge();
				ws.Range["A7"].Value = "Принятые";
				ws.Range["A7:A8"].Merge();
				ws.Range["A9"].Value = "Непринятые";
				ws.Range["A9:A18"].Merge();
				ws.Range["A19"].Value = "Ошибочные";
				ws.Range["A19:A20"].Merge();
				ws.Range["A21"].Value = "Набранные";
				ws.Range["A21:A22"].Merge();
				ws.Range["A1:A22"].VerticalAlignment = XlHAlign.xlHAlignCenter;
				ws.Range["A5:A22"].Borders.LineStyle = XlLineStyle.xlDot;
				ws.Range["A5:A22"].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
				ws.Columns[1].ColumnWidth = 14;

				string[] header = {
					"Имя файла",
					"Имя рабочей станции",
					"Дата создания",
					"Отчетный период",
					"Количество",
					"Длительность",
					"Количество",
					"Длительность",
					"Количество",
					"Длительность",
					"% от всего",
					"Регистратура перезвонила и дозвонилась",
					"Пациент перезвонил самостоятельно",
					"Не перезвонили / не дозвонились",
					"% недозвона",
					"Регамент соблюден",
					"Регламент нарушен",
					"% нарушения регламента",
					"Количество",
					"Время",
					"Количество",
					"Время",
					"Расположение" };

				int rowToFillHead = 1;
				foreach (string name in header) {
					ws.Range["B" + rowToFillHead].Value = name;
					rowToFillHead++;
				}
				ws.Range["B1:B23"].Borders.LineStyle = XlLineStyle.xlDot;
				ws.Range["B1:B23"].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
				ws.Columns[2].ColumnWidth = 40;

				int columnToFill = 3;
				foreach(KeyValuePair<string, SpRecordFileInformation> fileInfo in filesInfo) {
					SpRecordFileInformation info = fileInfo.Value;

					double missedPercentage = (double)info.callsMissed / (double)info.callsTotal;
					double callsBackNotPercentage = (double)info.callsBackNot / (double)info.callsMissed;
					double regulatinNotObservedPercentage = (double)info.missedCallsRegulationNotObserved /
						((double)info.missedCallsRegulationNotObserved + (double)info.missedCallsRegulationObserved);

					string[] values = {
						Path.GetFileName(fileInfo.Key),
						info.workstationName,
						info.creationDate,
						info.accountingPeriod,
						info.callsTotal.ToString(),
						StringFromTimeSpan(info.timeTotal),
						info.callsAccepted.ToString(),
						StringFromTimeSpan(info.timeAccepted),
						info.callsMissed.ToString(),
						StringFromTimeSpan(info.timeMissed),
						string.Format("{0:P2}", missedPercentage),
						info.callsBackByRegistry.ToString(),
						info.callsBackByPatient.ToString(),
						info.callsBackNot.ToString(),
						string.Format("{0:P2}", callsBackNotPercentage),
						info.missedCallsRegulationObserved.ToString(),
						info.missedCallsRegulationNotObserved.ToString(),
						string.Format("{0:P2}", regulatinNotObservedPercentage),
						info.callsAccidential.ToString(),
						StringFromTimeSpan(info.timeAccidential),
						info.callsDialed.ToString(),
						StringFromTimeSpan(info.timeDialed),
						"Лист " + (columnToFill - 2)};

					int rowToFillValues = 1;

					string columnName = GetExcelColumnName(columnToFill);
					foreach (string value in values) {
						ws.Range[columnName + rowToFillValues].Value = value;
						rowToFillValues++;
					}

					ws.Columns[columnToFill].ColumnWidth = 25;
					ws.Range[columnName + "1:" + columnName + (rowToFillValues - 1)].
						Borders.LineStyle = XlLineStyle.xlDot;
					ws.Range[columnName + "1:" + columnName + (rowToFillValues - 1)].
						BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
					columnToFill++;
				}

				ws.Rows[2].Font.Bold = true;
				ws.Columns[1].Font.Bold = true;

				string lastUsedColumn = GetExcelColumnName(columnToFill - 1);
				string[] rangesBorderAround = {
					"A5:" + lastUsedColumn + "6",
					"A7:" + lastUsedColumn + "8",
					"B9:" + lastUsedColumn + "11",
					"B12:" + lastUsedColumn + "15",
					"B16:" + lastUsedColumn + "18",
					"A19:" + lastUsedColumn + "20",
					"A21:" + lastUsedColumn + "22",
					"B23:" + lastUsedColumn + "23" };

				foreach (string rangeBorderAround in rangesBorderAround) {
					ws.Range[rangeBorderAround].
						BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
				}

				string[] rangesInteriorColor = {
					"B11:" + lastUsedColumn + "11",
					"B15:" + lastUsedColumn + "15",
					"B18:" + lastUsedColumn + "18" };

				foreach (string rangeInteriorColor in rangesInteriorColor) {
					ws.Range[rangeInteriorColor].Interior.ColorIndex = 36;
				}

				ws.Range["B2:" + lastUsedColumn + "2"].Interior.ColorIndex = 24;
				ws.Application.ActiveWindow.SplitColumn = 2;
				ws.Application.ActiveWindow.FreezePanes = true;
				
				string pathToSave = Directory.GetCurrentDirectory() + "\\" +
					"Анализ звонков - " + DateTime.Now.ToLocalTime().ToString().Replace(":", ".") + ".xlsx";
				wb.SaveAs(pathToSave);

				UpdateTextBox("Книга с отчетом сохранена по адресу: " + pathToSave);
				LoggingSystem.LogMessageToFile("Анализ звонков сохранен в файл: " + pathToSave);

				wb.Close(SaveChanges: false);
				xlApp.Quit();

				Process.Start(pathToSave);

				return true;
			} catch (Exception e) {
				LoggingSystem.LogMessageToFile(e.Message);
				LoggingSystem.LogMessageToFile(e.StackTrace);
			}

			return false;
		}

		private static string StringFromTimeSpan(TimeSpan timeSpan) {
			return (int)timeSpan.TotalHours + ":" + timeSpan.Minutes + ":" + timeSpan.Seconds;
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

		private void UpdateProgressBar(int percentage) {
			if (progressBar == null)
				return;
			if (percentage > 100)
				percentage = 100;

			progressBar.BeginInvoke((MethodInvoker)delegate {
				progressBar.Value = percentage;
			});
		}

		private void UpdateTextBox(string message, bool newSection = false, bool error = false) {
			if (textBox == null) return;
			textBox.BeginInvoke((MethodInvoker)delegate {
				if (newSection)
					textBox.AppendText("-------------------------------" +
						Environment.NewLine);

				if (error)
					textBox.AppendText("===== ВНИМАНИЕ! ОШИБКА! =====" +
						Environment.NewLine);

				textBox.AppendText(DateTime.Now.ToString("HH:mm:ss") + ": " +
					message + Environment.NewLine);
			});

			LoggingSystem.LogMessageToFile(message);
		}
	}
}
