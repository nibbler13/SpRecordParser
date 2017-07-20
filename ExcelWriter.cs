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

		public ExcelWriter(
			ProgressBar progressBar, 
			System.Windows.Forms.TextBox textBox,
			Dictionary<string, SpRecordFileInformation> filesInfo) {
			this.progressBar = progressBar;
			this.textBox = textBox;
			this.filesInfo = filesInfo;
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

				float initialPercent = progressBar.Value;
				float oneFilePercentage = (90.0f - initialPercent) / (float)filesInfo.Count;

				foreach(KeyValuePair<string, SpRecordFileInformation> fileInfo in filesInfo) {
					List<List<string>> fileContent = fileInfo.Value.fileContent;
					UpdateTextBox("Заполнение информации по файлу: " + fileInfo.Key);

					float initialFilePercent = progressBar.Value;
					float oneLinePercent = oneFilePercentage / fileContent.Count;

					int rowToFill = 1;
					foreach(List<string> line in fileContent) {
						initialFilePercent += oneLinePercent;
						UpdateProgressBar((int)initialFilePercent);

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
					initialPercent += oneFilePercentage;
					UpdateProgressBar((int)initialPercent);
				}

				UpdateTextBox("Создание сводной таблицы по всем файлам");

				string firstColumnRange = "A5:A34";
				Dictionary<string, string> firstColumn = new Dictionary<string, string> {
					{"Всего звонков", "A5:A6" },
					{"Принятые", "A7:A9" },
					{"Непринятые", "A10:A29" },
					{"Ошибочные" + Environment.NewLine + "непринятые", "A30:A33" },
					{"Набранные", "A34:A35" }
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
					"Первая попытка - не позднее " + Properties.Settings.Default.CallbackFirstAttemptMax +
					" минут после непринятого вызова" + Environment.NewLine +
					"Вторая попытка - от " + Properties.Settings.Default.CallbackFirstAttemptMax +
					" до " + Properties.Settings.Default.CallbackSecondAttemptMax +
					" минут после непринятого вызова" + Environment.NewLine +
					"Третья попытка - от " + Properties.Settings.Default.CallbackSecondAttemptMax +
					" до " + Properties.Settings.Default.CallbackThirdAttemptMax + " минут после непринятого вызова",
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
					"Длительность меньше 6 секунд (количество)",
					"Некорректные данные" + (Properties.Settings.Default.CalcRepeatedMissedAsOne ? ", повторяющиеся с одним номером" : "") +
					(Properties.Settings.Default.IgnoreInternalMissedCalls ? ", внутренние номера" : "") +
					(Properties.Settings.Default.IgnoreNonworkingTimeMissedCalls ? ", нерабочие часы с " +
					Properties.Settings.Default.WorkingTimeBegin.ToString() + " до " + 
					Properties.Settings.Default.WorkingTimeEnd.ToString() : "") + " (количество)",
					"Общая длительность",
					//dialed
					"% от всех входящих",
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

				string headersRange = "B1:C36";
				ws.Range[headersRange].Borders.LineStyle = XlLineStyle.xlDot;
				ws.Range[headersRange].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
				ws.Range[headersRange].VerticalAlignment = XlHAlign.xlHAlignCenter;
				//ws.Range[headersRange].HorizontalAlignment = XlVAlign.xlVAlignCenter;
				ws.Rows[14].RowHeight = 60;
				ws.Columns[2].ColumnWidth = 42;
				ws.Columns[3].ColumnWidth = 20;

				int columnToStartFill = 4;
				int listStartPosition = 1;
				foreach (KeyValuePair<string, SpRecordFileInformation> fileInfo in filesInfo) {
					SpRecordFileInformation info = fileInfo.Value;

					int callsIncoming = info.CallsAccepted + info.CallsMissed + 
						info.CallsMissedAccidentialShort + info.CallsMissedAccidentialWrongValues;
					int ringUpFailed = info.RingUpDidNotTried + 
									   info.RingUpNotRegulationObserved + 
									   info.RingUpNotRegulationNotObserved;
					int regulationNotObserved = info.RingUp1tryRegulationNotObserved +
												info.RingUp2tryRegulationNotObserved +
												info.RingUp3tryRegulationNotObserved +
												info.RingUp3MoreTryRegulationNotObserved +
												info.RingUpByPatientRegulationNotObserved +
												info.RingUpNotRegulationNotObserved +
												info.RingUpDidNotTried;
					int callsAccidential = info.CallsMissedAccidentialShort + info.CallsMissedAccidentialWrongValues;

					string[] values = {
						Path.GetFileName(fileInfo.Key),				//Имя файла
						info.WorkstationName,						//Имя рабочей станции
						info.CreationDate,							//Дата создания
						info.AccountingPeriod,						//Отчетный период
						info.CallsTotal.ToString(),					//Количество
						StringFromTimeSpan(info.TimeTotal),			//Общая длительность
						info.CallsAccepted.ToString(),				//Количество
						StringFromTimeSpan(info.TimeAccepted),		//Общая длительность
						StringFromTimeSpan(GetAverageDuration(info.TimeAccepted, info.CallsAccepted)), //Средняя длительность
						info.CallsMissed.ToString(),				//Количество
						StringFromTimeSpan(info.TimeMissed),		//Общая длительность
						StringFromTimeSpan(GetAverageDuration(info.TimeMissed, info.CallsMissed)), //Средняя длительность
						GetPercentage(info.CallsMissed, callsIncoming),	//% от всех входящих
						"",											//Регламент работы
						info.RingUp1tryRegulationObserved.ToString(), //Дозвонились с одной попытки Регламент соблюден
						info.RingUp1tryRegulationNotObserved.ToString(), //Дозвонились с одной попытки Регламент нарушен
						info.RingUp2tryRegulationObserved.ToString(), //Дозвонились с двух попыток Регламент соблюден
						info.RingUp2tryRegulationNotObserved.ToString(), //Дозвонились с двух попыток Регламент нарушен
						info.RingUp3tryRegulationObserved.ToString(), //Дозвонились с трех попыток Регламент соблюден
						info.RingUp3tryRegulationNotObserved.ToString(), //Дозвонились с трех попыток Регламент нарушен
						info.RingUp3MoreTryRegulationObserved.ToString(), //Дозвонились с более чем трех попыток Регламент соблюден
						info.RingUp3MoreTryRegulationNotObserved.ToString(), //Дозвонились с более чем трех попыток Регламент нарушен
						info.RingUpByPatientRegulationObserved.ToString(), //Пациент перезвонил самостоятельно Регламент соблюден
						info.RingUpByPatientRegulationNotObserved.ToString(), //Пациент перезвонил самостоятельно Регламент нарушен
						info.RingUpNotRegulationObserved.ToString(), //Не дозвонились Регламент соблюден
						info.RingUpNotRegulationNotObserved.ToString(), //Не дозвонились Регламент нарушен
						info.RingUpDidNotTried.ToString(), //Не пытались перезвонить Регламент нарушен
						GetPercentage(ringUpFailed, info.CallsMissed), //% недозвона
						GetPercentage(regulationNotObserved, info.CallsMissed), //% нарушения регламента
						info.CallsMissedAccidentialShort.ToString(),			//Длительность меньше 6 секунд (количество)
						info.CallsMissedAccidentialWrongValues.ToString(),		//Неправильные данные (количество)
						StringFromTimeSpan(info.TimeAccidential),	//Общая длительность
						GetPercentage(callsAccidential, callsIncoming), //% от всех входящих
						info.CallsDialed.ToString(),				//Количество
						StringFromTimeSpan(info.TimeDialed),		//Общая длительность
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
						double missedPercetage = (double)info.CallsMissed / (double)callsIncoming * 100;
						if (missedPercetage <= Properties.Settings.Default.MissedCallsGoodMax) {
							colorIndex = 36; //green
						} else if (missedPercetage > Properties.Settings.Default.MissedCallsBadMin) {
							colorIndex = 3; //red
						} else {
							colorIndex = 6; //yellow
						}

						if (colorIndex > 0) {
							ws.Range[columnName + "13"].Interior.ColorIndex = colorIndex;
							colorIndex = 0;
						}
					}

					if (info.CallsMissed > 0) {
						double regulationsNotObservedPercentage =
							(double)regulationNotObserved / (double)info.CallsMissed * 100;
						if (regulationsNotObservedPercentage <= Properties.Settings.Default.RegulationGoodMax) {
							colorIndex = 36; //green
						} else if (regulationsNotObservedPercentage > Properties.Settings.Default.RegulationBadMin) {
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
					"A30:" + lastUsedColumn + "33",
					//dialed
					"A34:" + lastUsedColumn + "35",
					//position
					"B36:" + lastUsedColumn + "36" };

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

				ws.Rows[30].RowHeight = 15;
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
