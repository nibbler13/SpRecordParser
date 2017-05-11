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
						//			  = 12 - missed
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

				string firstColumnRange = "A5:A34";
				Dictionary<string, string> firstColumn = new Dictionary<string, string> {
					{"Всего звонков", "A5:A6" },
					{"Принятые", "A7:A9" },
					{"Непринятые", "A10:A29" },
					{"Ошибочные", "A30:A32" },
					{"Набранные", "A33:A34" }
				};

				foreach (KeyValuePair<string, string> item in firstColumn) {
					try {
						ws.Range[item.Value.Split(':')[0]].Value = item.Key;
						ws.Range[item.Value].Merge();
					} catch (Exception e) {
						LoggingSystem.LogMessageToFile(e.Message);
						LoggingSystem.LogMessageToFile(e.StackTrace);
					}
				}

				ws.Range[firstColumnRange].VerticalAlignment = XlHAlign.xlHAlignCenter;
				ws.Range[firstColumnRange].HorizontalAlignment = XlVAlign.xlVAlignCenter;
				ws.Range[firstColumnRange].Borders.LineStyle = XlLineStyle.xlDot;
				ws.Range[firstColumnRange].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
				ws.Columns[1].ColumnWidth = 13;

				string[] secondColumn = {
					//header
					"Имя файла",
					"Имя рабочей станции",
					"Дата создания",
					"Отчетный период",
					//total
					"Количество",
					"Общая длительность",
					//accepted
					"Количество",
					"Общая длительность",
					"Средняя длительность",
					//missed
					"Количество",
					"Общая длительность",
					"Средняя длительность",
					"% от всех входящих",
					"Регламент работы с непринятыми вызовами:" + Environment.NewLine +
					"Первая попытка - не позднее 5 минут после непринятого вызова" + Environment.NewLine +
					"Вторая попытка - от 5 до 20 минут после непринятого вызова" + Environment.NewLine +
					"Третья попытка - от 20 до 35 минут после непринятого вызова",
					"Дозвонились с одной попытки",
					"Дозвонились с одной попытки",
					"Дозвонились с двух попыток",
					"Дозвонились с двух попыток",
					"Дозвонились с трех попыток",
					"Дозвонились с трех попыток",
					"Дозвонились с более чем трех попыток",
					"Дозвонились с более чем трех попыток",
					"Пациент перезвонил самостоятельно",
					"Пациент перезвонил самостоятельно",
					"Не дозвонились",
					"Не дозвонились",
					"Не пытались перезвонить",
					"% недозвона",
					"% нарушения регламента",
					//accidental
					"Количество",
					"Общая длительность",
					"% от всех входящих",
					//dialed
					"Количество",
					"Общая длительность",
					//header
					"Расположение в книге" };

				int rowToStartFill = 1;
				for (int i = 0; i < secondColumn.Length; i++) {
					ws.Range["B" + rowToStartFill].Value = secondColumn[i];

					if (i < secondColumn.Length - 1 && secondColumn[i] == secondColumn[i + 1]) {
						ws.Range["B" + rowToStartFill + ":B" + (rowToStartFill + 1)].Merge();
						i++;
						rowToStartFill++;
					} else {
						if (!secondColumn[i].Equals("Не пытались перезвонить"))
							ws.Range["B" + rowToStartFill + ":C" + rowToStartFill].Merge();
					}

					rowToStartFill++;
				}

				string[] thirdColumn = {
					"Регламент соблюден",
					"Регламент нарушен",
					"Регламент соблюден",
					"Регламент нарушен",
					"Регламент соблюден",
					"Регламент нарушен",
					"Регламент соблюден",
					"Регламент нарушен",
					"Регламент соблюден",
					"Регламент нарушен",
					"Регламент соблюден",
					"Регламент нарушен",
					"Регламент нарушен"
				};

				rowToStartFill = 15;
				foreach (string item in thirdColumn) {
					ws.Range["C" + rowToStartFill].Value = item;
					rowToStartFill++;
				}

				string headersRange = "B1:C35";
				ws.Range[headersRange].Borders.LineStyle = XlLineStyle.xlDot;
				ws.Range[headersRange].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
				ws.Range[headersRange].VerticalAlignment = XlHAlign.xlHAlignCenter;
				//ws.Range[headersRange].HorizontalAlignment = XlVAlign.xlVAlignCenter;
				ws.Rows[14].RowHeight = 60;
				ws.Columns[2].ColumnWidth = 39;
				ws.Columns[3].ColumnWidth = 20;

				int columnToStartFill = 4;
				int listStartPosition = 1;
				foreach (KeyValuePair<string, SpRecordFileInformation> fileInfo in filesInfo) {
					SpRecordFileInformation info = fileInfo.Value;

					int callsIncoming = info.callsTotal - info.callsDialed;
					int ringUpFailed = info.ringUpDidNotTried + 
									   info.ringUpNotRegulationObserved + 
									   info.ringUpNotRegulationNotObserved;
					int regulationNotObserved = info.ringUp1tryRegulationNotObserved +
												info.ringUp2tryRegulationNotObserved +
												info.ringUp3tryRegulationNotObserved +
												info.ringUp3MoreTryRegulationNotObserved +
												info.ringUpByPatientRegulationNotObserved +
												info.ringUpNotRegulationNotObserved +
												info.ringUpDidNotTried;

					string[] values = {
						Path.GetFileName(fileInfo.Key),				//Имя файла
						info.workstationName,						//Имя рабочей станции
						info.creationDate,							//Дата создания
						info.accountingPeriod,						//Отчетный период
						info.callsTotal.ToString(),					//Количество
						StringFromTimeSpan(info.timeTotal),			//Общая длительность
						info.callsAccepted.ToString(),				//Количество
						StringFromTimeSpan(info.timeAccepted),		//Общая длительность
						StringFromTimeSpan(GetAverageDuration(info.timeAccepted, info.callsAccepted)), //Средняя длительность
						info.callsMissed.ToString(),				//Количество
						StringFromTimeSpan(info.timeMissed),		//Общая длительность
						StringFromTimeSpan(GetAverageDuration(info.timeMissed, info.callsMissed)), //Средняя длительность
						GetPercentage(info.callsMissed, callsIncoming),	//% от всех входящих
						"",											//Регламент работы
						info.ringUp1tryRegulationObserved.ToString(), //Дозвонились с одной попытки Регламент соблюден
						info.ringUp1tryRegulationNotObserved.ToString(), //Дозвонились с одной попытки Регламент нарушен
						info.ringUp2tryRegulationObserved.ToString(), //Дозвонились с двух попыток Регламент соблюден
						info.ringUp2tryRegulationNotObserved.ToString(), //Дозвонились с двух попыток Регламент нарушен
						info.ringUp3tryRegulationObserved.ToString(), //Дозвонились с трех попыток Регламент соблюден
						info.ringUp3tryRegulationNotObserved.ToString(), //Дозвонились с трех попыток Регламент нарушен
						info.ringUp3MoreTryRegulationObserved.ToString(), //Дозвонились с более чем трех попыток Регламент соблюден
						info.ringUp3MoreTryRegulationNotObserved.ToString(), //Дозвонились с более чем трех попыток Регламент нарушен
						info.ringUpByPatientRegulationObserved.ToString(), //Пациент перезвонил самостоятельно Регламент соблюден
						info.ringUpByPatientRegulationNotObserved.ToString(), //Пациент перезвонил самостоятельно Регламент нарушен
						info.ringUpNotRegulationObserved.ToString(), //Не дозвонились Регламент соблюден
						info.ringUpNotRegulationNotObserved.ToString(), //Не дозвонились Регламент нарушен
						info.ringUpDidNotTried.ToString(), //Не пытались перезвонить Регламент нарушен
						GetPercentage(ringUpFailed, info.callsMissed), //% недозвона
						GetPercentage(regulationNotObserved, info.callsMissed), //% нарушения регламента
						info.callsAccidential.ToString(),			//Количество
						StringFromTimeSpan(info.timeAccidential),	//Общая длительность
						GetPercentage(info.callsAccidential, callsIncoming), //% от всех входящих
						info.callsDialed.ToString(),				//Количество
						StringFromTimeSpan(info.timeDialed),		//Общая длительность
						"Лист " + listStartPosition};			//Расположение

					int rowToFillValues = 1;
					listStartPosition++;

					string columnName = GetExcelColumnName(columnToStartFill);
					foreach (string value in values) {
						ws.Range[columnName + rowToFillValues].Value = value;
						rowToFillValues++;
					}

					int colorIndex = 0;
					if (callsIncoming > 0) {
						double missedPercetage = (double)info.callsMissed / (double)callsIncoming * 100;
						if (missedPercetage <= 4.0) {
							colorIndex = 36; //green
						} else if (missedPercetage > 6.0) {
							colorIndex = 3; //red
						} else {
							colorIndex = 6; //yellow
						}

						if (colorIndex > 0) {
							ws.Range[columnName + "13"].Interior.ColorIndex = colorIndex;
							colorIndex = 0;
						}
					}

					if (info.callsMissed > 0) {
						double regulationsNotObservedPercentage =
							(double)regulationNotObserved / (double)info.callsMissed * 100;
						if (regulationsNotObservedPercentage <= 5.0) {
							colorIndex = 36; //green
						} else if (regulationsNotObservedPercentage > 15.0) {
							colorIndex = 3; //red
						} else {
							colorIndex = 6; //yellow
						}

						if (colorIndex > 0) {
							ws.Range[columnName + "29"].Interior.ColorIndex = colorIndex;
						}
					}


					ws.Columns[columnToStartFill].ColumnWidth = 25;
					ws.Range[columnName + "1:" + columnName + (rowToFillValues - 1)].
						Borders.LineStyle = XlLineStyle.xlDot;
					ws.Range[columnName + "1:" + columnName + (rowToFillValues - 1)].
						BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
					columnToStartFill++;
				}

				ws.Rows[2].Font.Bold = true;
				ws.Columns[1].Font.Bold = true;

				string lastUsedColumn = GetExcelColumnName(columnToStartFill - 1);
				string[] rangesBorderAround = {
					//total
					"A5:" + lastUsedColumn + "6",
					//accepted
					"A7:" + lastUsedColumn + "9",
					//missed
					//at glance
					"B10:" + lastUsedColumn + "13",
					//regulation
					"B14:" + lastUsedColumn + "27",
					//result
					"B28:" + lastUsedColumn + "29",
					//accidential
					"A30:" + lastUsedColumn + "32",
					//dialed
					"A33:" + lastUsedColumn + "34",
					//position
					"B35:" + lastUsedColumn + "35" };

				foreach (string rangeBorderAround in rangesBorderAround) {
					ws.Range[rangeBorderAround].
						BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
				}

				string[] rangesInteriorColor = {
					"B13:C13",
					"B29:C29" };

				foreach (string rangeInteriorColor in rangesInteriorColor) {
					ws.Range[rangeInteriorColor].Interior.ColorIndex = 36;
				}

				ws.Range["B2:" + lastUsedColumn + "2"].Interior.ColorIndex = 24;
				ws.Application.ActiveWindow.SplitColumn = 3;
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

		private TimeSpan GetAverageDuration(TimeSpan duration, int quantity) {
			TimeSpan averageDuration = new TimeSpan();

			if (quantity != 0)
				averageDuration = new TimeSpan(duration.Ticks / quantity);

			return averageDuration;
		}

		private string GetPercentage(int value, int total) {
			string percentage = "Значение недоступно";

			if (total != 0)
				percentage = string.Format("{0:P2}", (double)value / (double)total);

			return percentage;
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
