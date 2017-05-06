using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace SpRecordParser {
	class ExcelWriter {
		public static bool WriteToExcel(Dictionary<string, SpRecordFileInformation> filesInfo) {
			try {
				Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

				if (xlApp == null) {
					LoggingSystem.LogMessageToFile("Не удалось запустить MS Excel");
					return false;
				}

				xlApp.Visible = true;
				//xlApp.EnableAnimations = false;
				xlApp.DisplayAlerts = false;

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

				foreach(KeyValuePair<string, SpRecordFileInformation> fileInfo in filesInfo) {
					List<List<string>> fileContent = fileInfo.Value.fileContent;

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
				}

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
					"Время" };

				int rowToFillHead = 1;
				foreach (string name in header) {
					ws.Range["B" + rowToFillHead].Value = name;
					rowToFillHead++;
				}
				ws.Columns[2].ColumnWidth = 40;

				int columnToFill = 3;
				foreach(KeyValuePair<string, SpRecordFileInformation> fileInfo in filesInfo) {
					SpRecordFileInformation info = fileInfo.Value;

					string[] values = {
						fileInfo.Key,
						info.workstationName,
						info.creationDate,
						info.accountingPeriod,
						info.callsTotal.ToString(),
						StringFromTimeSpan(info.timeTotal),
						info.callsAccepted.ToString(),
						StringFromTimeSpan(info.timeAccepted),
						info.callsMissed.ToString(),
						StringFromTimeSpan(info.timeMissed),
						((double)info.callsMissed / 
						(double)info.callsTotal).ToString(),
						info.callsBackByRegistry.ToString(),
						info.callsBackByPatient.ToString(),
						info.callsBackNot.ToString(),
						((double)info.callsBackNot / 
						(double)info.callsMissed).ToString(),
						info.missedCallsRegulationObserved.ToString(),
						info.missedCallsRegulationNotObserved.ToString(),
						((double)info.missedCallsRegulationNotObserved / 
						((double)info.missedCallsRegulationNotObserved +
						(double)info.missedCallsRegulationObserved)).ToString(),
						info.callsAccidential.ToString(),
						StringFromTimeSpan(info.timeAccidential),
						info.callsDialed.ToString(),
						StringFromTimeSpan(info.timeDialed)};

					int rowToFillValues = 1;

					foreach (string value in values) {
						ws.Range[GetExcelColumnName(columnToFill) + rowToFillValues].Value = value;
						rowToFillValues++;
					}

					ws.Columns[columnToFill].ColumnWidth = 25;
					columnToFill++;
				}

				ws.Rows[2].Font.Bold = true;
				ws.Columns[1].Font.Bold = true;
				string lastUsedColumn = GetExcelColumnName(columnToFill - 1);
				ws.UsedRange.Borders.LineStyle = XlLineStyle.xlContinuous;


				//string pathToSave = Directory.GetCurrentDirectory() + "\\" +
				//	"Расчет стоимости - " + Path.GetFileName(filePath) + " " +
				//	DateTime.Now.ToLocalTime().ToString().Replace(":", ".") + ".xlsx";
				//Console.WriteLine("---- " + pathToSave);
				//wb.SaveAs(pathToSave);

				//LoggingSystem.LogMessageToFile("Результат расчета сохранен в файл: " + pathToSave);

				//xlApp.EnableAnimations = true;
				xlApp.DisplayAlerts = true;
				xlApp.Quit();

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
	}
}
