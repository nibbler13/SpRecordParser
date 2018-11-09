using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using System.Globalization;

namespace SpRecordParser {
	class FileParser {
		private ProgressBar progressBar;
		private TextBox textBox;
		private Dictionary<string, SpRecordFileInformation> filesInfo;
		private int percentageCompleted;

		public FileParser(ProgressBar progressBar, TextBox textBox) {
			this.progressBar = progressBar;
			this.textBox = textBox;
			filesInfo = new Dictionary<string, SpRecordFileInformation>();
			percentageCompleted = 0;
		}

		public void ParseFiles(List<string> fileNames) {
			UpdateTextBox("Начало анализа");

			int oneFileProgress = 20 / fileNames.Count;
			foreach (string fileName in fileNames) {
				UpdateTextBox("Файл:" + fileName, newSection: true);
				if (!IsFileExistAndNotEmpty(fileName)) {
					UpdateTextBox("Файл не существует или пустой", error: true);
					continue;
				}

				List<List<string>> fileContent = GetCsvFileContent(fileName);

				if (fileContent.Count == 0) {
					UpdateTextBox("Не удалось прочитать файл", error: true);
					continue;
				}

				AnalyseFileContentAndAddToDictionary(fileName, fileContent);
				percentageCompleted += oneFileProgress;
				UpdateProgressBar(percentageCompleted);
			}

			if (filesInfo.Count == 0) {
				MessageBox.Show(
					"Результирующий файл пуст", "Ошибка", 
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			UpdateProgressBar(20);
			UpdateTextBox("Выгрузка данных в Excel", newSection: true);

			ExcelWriter excelWriter = new ExcelWriter(progressBar, textBox, filesInfo);
			if (!excelWriter.WriteToExcel()) {
				MessageBox.Show(
					"Обработка данных завершена с ошибкой", "SpRecordParser", 
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			UpdateProgressBar(100);

			MessageBox.Show(
				"Обработка данных завершена успешно", "SpRecordParser",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void AnalyseFileContentAndAddToDictionary(string fileName, List<List<string>> fileContent) {
			if (filesInfo.ContainsKey(fileName)) {
				UpdateTextBox("Файл уже проанализирован ранее");
				return;
			}

			if (fileContent.Count < 7) {
				UpdateTextBox("Файл не соответствует формату SpRecord" +
					"(должно быть минимум 7 строк, в файле: " + 
					fileContent.Count, error: true);
				return; 
			}

			List<string> lastRow = fileContent.Last();
			if (lastRow.Count < 1)
				UpdateTextBox("Не соответсвует формат последней строки, " +
					"должен быть хотя бы 1 столбец", error: true);

			string lastRowText = lastRow[0];
			if (!lastRowText.StartsWith("Всего записей:")) {
				UpdateTextBox("В последней строке отсутствует информация об общем " +
					"количестве записей");
				return;
			}

			SpRecordFileInformation fileInformation = ParseLastRow(lastRowText);

			for (int row = fileContent.Count - 2; row >= 0; row--) {
				List<string> line = fileContent[row];

				if (line.Count < 8 && line.Count != 1) {
					UpdateTextBox("Размер строки " + (row + 1) + " не совпадает с форматом SpRecord");
					continue;
				}

				if (line.Count == 1) {
					string text = line[0];

					if (text.StartsWith("Выбраны записи"))
						ParseLineAccountingPeriod(text, ref fileInformation);

					if (text.StartsWith("Рабочая станция"))
						ParseLineWorkstation(text, ref fileInformation);

					if (text.StartsWith("Отчет создан"))
						ParseLineCreationDate(text, ref fileInformation);
				}

				if (line.Count >= 8) {
					if (line[0].Equals("Название канала"))
						continue;

					try {
						DateTime dateTime = DateTime.Parse(line[1]);
						dateTime = dateTime.AddMilliseconds(dateTime.TimeOfDay.TotalMilliseconds * -1);

						if (!fileInformation.DaysInfo.ContainsKey(dateTime))
							fileInformation.DaysInfo.Add(dateTime, new SpRecordFileInformation.DayInfo());

						SpRecordFileInformation.DayInfo dayInfo = fileInformation.DaysInfo[dateTime];

						TimeSpan duration = ParseTimeSpan(line[2]);
						double totalSeconds = duration.TotalSeconds;
						string phoneNumbers = line[4];

						if (phoneNumbers.EndsWith("-> 601") || phoneNumbers.EndsWith("-> 611")) {
							dayInfo.TotalIncoming++;

							if (totalSeconds <= 5) {
								dayInfo.IncrementMissedCallCount(
									SpRecordFileInformation.DayInfo.MissedCallType.ConditionalyLostSelf5sec);
							} else if (totalSeconds > 5 && totalSeconds <= 10) {
								dayInfo.IncrementMissedCallCount(
									SpRecordFileInformation.DayInfo.MissedCallType.ConditionalyLostSelf10sec);
							} else if (totalSeconds > 10 && totalSeconds <= 15) {
								dayInfo.IncrementMissedCallCount(
									SpRecordFileInformation.DayInfo.MissedCallType.ConditionalyLostSelf15sec);
							} else if (totalSeconds > 15 && totalSeconds <= 20) {
								dayInfo.IncrementMissedCallCount(
									SpRecordFileInformation.DayInfo.MissedCallType.ConditionalyLostSelf20sec);
							} else if (totalSeconds > 20 && totalSeconds <= 25) {
								dayInfo.IncrementMissedCallCount(
									SpRecordFileInformation.DayInfo.MissedCallType.ConditionalyLostSelf25sec);
							} else if (totalSeconds > 25 && totalSeconds <= 30) {
								dayInfo.IncrementMissedCallCount(
									SpRecordFileInformation.DayInfo.MissedCallType.ConditionalyLostSelf30sec);
							}
						} else if (phoneNumbers.EndsWith("-> 30400")) {
							dayInfo.TotalRedirected++;

							if (totalSeconds <= 5) {
								dayInfo.IncrementMissedCallCount(
									SpRecordFileInformation.DayInfo.MissedCallType.ConditionalyLostRedirected5sec);
							} else if (totalSeconds > 5 && totalSeconds <= 10) {
								dayInfo.IncrementMissedCallCount(
									SpRecordFileInformation.DayInfo.MissedCallType.ConditionalyLostRedirected10sec);
							} else if (totalSeconds > 10 && totalSeconds <= 15) {
								dayInfo.IncrementMissedCallCount(
									SpRecordFileInformation.DayInfo.MissedCallType.ConditionalyLostRedirected15sec);
							} else if (totalSeconds > 15 && totalSeconds <= 20) {
								dayInfo.IncrementMissedCallCount(
									SpRecordFileInformation.DayInfo.MissedCallType.ConditionalyLostRedirected20sec);
							} else if (totalSeconds > 20 && totalSeconds <= 25) {
								dayInfo.IncrementMissedCallCount(
									SpRecordFileInformation.DayInfo.MissedCallType.ConditionalyLostRedirected25sec);
							} else if (totalSeconds > 25 && totalSeconds <= 30) {
								dayInfo.IncrementMissedCallCount(
									SpRecordFileInformation.DayInfo.MissedCallType.ConditionalyLostRedirected30sec);
							}
						}
					} catch (Exception e) {
						Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
					}

					//TimeSpan duration = ParseTimeSpan(line[2]);
					//string type = line[3];

					//if (type.Contains("Принятый")) {
					//	fileInformation.CallsAccepted++;
					//	fileInformation.TimeAccepted = fileInformation.TimeAccepted.Add(duration);
					//} else if (type.Contains("Набранный")) {
					//	fileInformation.CallsDialed++;
					//	fileInformation.TimeDialed = fileInformation.TimeDialed.Add(duration);
					//} else if (type.Contains("Непринятый")) {
					//	if (line.Count > 9) {
					//		fileInformation.CallsMissedAccidentialWrongValues++;
					//		fileInformation.TimeAccidential = fileInformation.TimeAccidential.Add(duration);
					//		continue;
					//	}

					//	if (duration.TotalSeconds <= 5) {
					//		fileInformation.CallsMissedAccidentialShort++;
					//		fileInformation.TimeAccidential = fileInformation.TimeAccidential.Add(duration);
					//		fileContent[row].Add("Ошибочный, длительность меньше 6 секунд");
					//	} else {
					//		if (IsAnalysingAMissedCallSucceed(row, ref fileInformation, ref fileContent)) {
					//			fileInformation.CallsMissed++;
					//			fileInformation.TimeMissed = fileInformation.TimeMissed.Add(duration);
					//		} else {
					//			fileInformation.CallsMissedAccidentialWrongValues++;
					//			fileInformation.TimeAccidential = fileInformation.TimeAccidential.Add(duration);
					//		}
					//	}
					//} else {
					//	UpdateTextBox("Неизвестный тип звонка: " + type);
					//}
				}
			}

			fileInformation.FileContent = fileContent;
			filesInfo.Add(fileName, fileInformation);
		}

		private bool IsAnalysingAMissedCallSucceed(
			int row, ref SpRecordFileInformation fileInformation, ref List<List<string>> fileContent) {
			//UpdateTextBox("Анализ пропущенного звонка, строка: " + (row + 1) + Environment.NewLine +
			//	string.Join(";", fileContent[row]));

			DateTime dateTimeMissedTime;
			if (!DateTime.TryParse(fileContent[row][1], out dateTimeMissedTime)) {
				fileContent[row].Add("Ошибочный, не удалось разобрать время звонка");
				UpdateTextBox("Не удалось разобрать время звонка, строка: " + (row + 1) +
					" значение: " + fileContent[row][1]);
				return false;
			} else if (Properties.Settings.Default.IgnoreNonworkingTimeMissedCalls &&
				(dateTimeMissedTime.TimeOfDay.TotalSeconds < Properties.Settings.Default.WorkingTimeBegin.TotalSeconds ||
				dateTimeMissedTime.TimeOfDay.TotalSeconds > Properties.Settings.Default.WorkingTimeEnd.TotalSeconds)) {
				fileContent[row].Add("Ошибочный, время звонка выходит за границы времени работы клиники");
				return false;
			}

			string[] phoneNumbers = SplitPhoneNumbers(fileContent[row][4]);
			string callerNumber = phoneNumbers[0];

			if (string.IsNullOrEmpty(callerNumber)) {
				fileContent[row].Add("Ошибочный, номер звонившего не удалось определить");
				UpdateTextBox("Номер звонившего не удалось определить");
				return false;
			} else if (callerNumber.Length <= 5 && Properties.Settings.Default.IgnoreInternalMissedCalls) {
				fileContent[row].Add("Ошибочный, внутренний номер");
				return false;
			}

			int callBackTries = 0;
			bool regulationsObserved = true;
			bool registryCallBackSucceded = false;
			bool conversationTookPlace = false;
			DateTime dateTimeCallBack = dateTimeMissedTime;

			for (int i = row - 1; i >= 0; i--) {
				if (fileContent[i].Count < 9)
					break;

				if (fileContent[i][0].Equals("Название канала"))
					break;

				DateTime dateTimeCurrentCall;
				if (!DateTime.TryParse(fileContent[i][1], out dateTimeCurrentCall)) {
					UpdateTextBox("Не удалось разобрать время звонка, строка: " + (i + 1) +
						" значение: " + fileContent[i][1]);
					continue;
				}

				if (!dateTimeMissedTime.Date.Equals(dateTimeCurrentCall.Date))
					break;

				string callPhoneNumbers = fileContent[i][4];
				if (!callPhoneNumbers.Contains(callerNumber))
					continue;

				dateTimeCallBack = dateTimeCurrentCall;

				string callType = fileContent[i][3];

				if (callType.Contains("Непринятый")) {
					if (!Properties.Settings.Default.CalcRepeatedMissedAsOne)
						break;

					String[] currentCallPhoneNumber = SplitPhoneNumbers(callPhoneNumbers);
					if (currentCallPhoneNumber[0].Length < 10)
						break;

					if (dateTimeCurrentCall.Subtract(dateTimeMissedTime).TotalSeconds >
						Properties.Settings.Default.CallbackThirdAttemptMax * 60)
						break;

					fileContent[i].Add("Ошибочный, дубль предыдущего непринятого звонка с таким же номером (" + (row + 1) + ")");
					continue;
				}

				fileContent[i].Add("Связка с пропущенным звонком");
				fileContent[i].Add("Строка: " + (row + 1));

				if (callType.Contains("Принятый")) {
					conversationTookPlace = true;
					break;
				} else if (callType.Contains("Набранный")) {
					string comment = fileContent[row][8];
					if (comment.Contains("Вызываемый абонент не ответил")) {
						callBackTries++;

						double minutesAfterMissedCall = dateTimeCurrentCall.Subtract(dateTimeMissedTime).TotalMinutes;
						CheckRegulationObservedStatus(ref regulationsObserved, callBackTries, minutesAfterMissedCall);
					} else {
						conversationTookPlace = true;
						registryCallBackSucceded = true;
						break;
					}
				}
			}

			string resultColumn1 = "";
			string resultColumn2 = "Регламент соблюден";
			string resultColumn3 = "";

			if (!conversationTookPlace) {
				resultColumn1 = "Не дозвонились";
				resultColumn3 = "Попыток сделано: " + callBackTries;

				if (callBackTries == 0) {
					regulationsObserved = false;
					fileInformation.RingUpDidNotTried++;
					resultColumn1 = "Не пытались перезвонить";
					resultColumn3 = "";
				} else if (callBackTries < 3) {
					regulationsObserved = false;
					fileInformation.RingUpNotRegulationNotObserved++;
				} else {
					if (regulationsObserved) {
						fileInformation.RingUpNotRegulationObserved++;
					} else {
						fileInformation.RingUpNotRegulationNotObserved++;
					}
				}
			} else {
				callBackTries++;
				double minutesAfterMissedCall = dateTimeCallBack.Subtract(dateTimeMissedTime).TotalMinutes;
				CheckRegulationObservedStatus(ref regulationsObserved, callBackTries, minutesAfterMissedCall);
				resultColumn3 = "Прошло минут: " + string.Format("{0:N2}", minutesAfterMissedCall) +
					", попыток сделано: " + callBackTries;

				if (registryCallBackSucceded) {
					switch (callBackTries) {
						case 1:
							resultColumn1 = "Дозвонились с одной попытки";
							if (regulationsObserved) {
								fileInformation.RingUp1tryRegulationObserved++;
							} else {
								fileInformation.RingUp1tryRegulationNotObserved++;
							}
							break;
						case 2:
							resultColumn1 = "Дозвонились с двух попыток";
							if (regulationsObserved) {
								fileInformation.RingUp2tryRegulationObserved++;
							} else {
								fileInformation.RingUp2tryRegulationNotObserved++;
							}
							break;
						case 3:
							resultColumn1 = "Дозвонились с трех попыток";
							if (regulationsObserved) {
								fileInformation.RingUp3tryRegulationObserved++;
							} else {
								fileInformation.RingUp3tryRegulationNotObserved++;
							}
							break;
						default:
							resultColumn1 = "Дозвонились с более чем трех попыток";
							if (regulationsObserved) {
								fileInformation.RingUp3MoreTryRegulationObserved++;
							} else {
								fileInformation.RingUp3MoreTryRegulationNotObserved++;
							}
							break;
					}
				} else {
					resultColumn1 = "Пациент перезвонил самостоятельно";
					if (regulationsObserved) {
						fileInformation.RingUpByPatientRegulationObserved++;
					} else {
						fileInformation.RingUpByPatientRegulationNotObserved++;
					}
				}
			}

			if (!regulationsObserved)
				resultColumn2 = "Регламент нарушен";

			fileContent[row].Add(resultColumn1);
			fileContent[row].Add(resultColumn2);
			fileContent[row].Add(resultColumn3);

			return true;
		}

		private void CheckRegulationObservedStatus(
			ref bool regulationsObserved, 
			int callBackTries, 
			double minutesAfterMissedCall) {
			switch (callBackTries) {
				case 1:
					if (minutesAfterMissedCall > Properties.Settings.Default.CallbackFirstAttemptMax)
						regulationsObserved = false;
					break;
				case 2:
					if (minutesAfterMissedCall <= Properties.Settings.Default.CallbackFirstAttemptMax ||
						minutesAfterMissedCall > Properties.Settings.Default.CallbackSecondAttemptMax)
						regulationsObserved = false;
					break;
				case 3:
					if (minutesAfterMissedCall <= Properties.Settings.Default.CallbackSecondAttemptMax || 
						minutesAfterMissedCall > Properties.Settings.Default.CallbackThirdAttemptMax)
						regulationsObserved = false;
					break;
				default:
					break;
			}

		}

		private string[] SplitPhoneNumbers(string str) {
			string[] phoneNumbers = new string[2];

			if (!str.Contains(" -> "))
				return phoneNumbers;

			try {
				phoneNumbers = str.Split(new[] { " -> " }, StringSplitOptions.None);

				int index = 0;
				foreach (string number in phoneNumbers) {
					string clearedToDigit = new string(number.Where(Char.IsDigit).ToArray());
					if (clearedToDigit.Length >= 10)
						clearedToDigit = clearedToDigit.Substring(clearedToDigit.Length - 10);
					if (!clearedToDigit.StartsWith("9") && clearedToDigit.Length >= 7)
						clearedToDigit = clearedToDigit.Substring(clearedToDigit.Length - 7);
					phoneNumbers[index] = clearedToDigit;
					index++;
				}
			} catch (Exception e) {
				LoggingSystem.LogMessageToFile(e.Message);
				LoggingSystem.LogMessageToFile(e.StackTrace);
				UpdateTextBox("Не удалось разобрать номера телефонов: " + str);
			}

			return phoneNumbers;
		}

		private TimeSpan ParseTimeSpan(string str) {
			TimeSpan timeSpan = new TimeSpan();

			try {
				string[] splitted = str.Split(':');
				timeSpan = new TimeSpan(
					int.Parse(splitted[0]), 
					int.Parse(splitted[1]), 
					int.Parse(splitted[2]));
			} catch (Exception e) {
				UpdateTextBox("Не удалось разобрать строку со временем: " + str +
					Environment.NewLine + "Ошибка: " + e.Message + " " + e.StackTrace);
			}

			return timeSpan;
		}

		private void ParseLineCreationDate(string str, ref SpRecordFileInformation fileInformation) {
			try {
				int colonSymbol = str.IndexOf(":");
				string creationDate = str.Substring(colonSymbol + 2, str.Length - colonSymbol - 3);
				fileInformation.CreationDate = creationDate;
				UpdateTextBox("Дата создания списка: " + creationDate);
			} catch (Exception e) {
				UpdateTextBox("Не удалось выполнить разбор строки с датой создания" + 
					Environment.NewLine + "Ошибка: " + e.Message + " " + e.StackTrace, error: true);
			}
		}

		private void ParseLineWorkstation(string str, ref SpRecordFileInformation fileInformation) {
			try {
				int colonSymbol = str.IndexOf(":");
				string workstationName = str.Substring(colonSymbol + 2, str.Length - colonSymbol - 3);
				fileInformation.WorkstationName = workstationName;
				UpdateTextBox("Рабочая станция: " + workstationName);
			} catch (Exception e) {
				UpdateTextBox("Не удалось выполнить разбор строки с рабочей станцией" +
					Environment.NewLine + "Ошибка: " + e.Message + " " + e.StackTrace, error: true);
			}
		}

		private void ParseLineAccountingPeriod(string str, ref SpRecordFileInformation fileInformation) {
			try {
				int mark = str.IndexOf("записи");
				string accountingPeriod = str.Substring(mark + 7, str.Length - mark - 7);
				fileInformation.AccountingPeriod = accountingPeriod;
				UpdateTextBox("Период отчета: " + accountingPeriod);
			} catch (Exception e) {
				UpdateTextBox("Не удалось выполнить разбор строки с периодом отчета" +
					Environment.NewLine + "Ошибка: " + e.Message + " " + e.StackTrace, error: true);
			}
		}

		private SpRecordFileInformation ParseLastRow(string str) {
			SpRecordFileInformation fileInformation = new SpRecordFileInformation();

			try {
				int colonSymbol = str.IndexOf(":");
				int dotSymbol = str.IndexOf('.');
				string totalRecordsText = str.Substring(
					colonSymbol + 2, dotSymbol - colonSymbol - 2);

				int totalRecords = 0;
				if (!int.TryParse(totalRecordsText, out totalRecords))
					UpdateTextBox("Не удалось считать общее количество записей", error: false);

				fileInformation.CallsTotal = totalRecords;

				int secondColonSymbol = str.IndexOf("ть:");
				string totalTimeText = str.Substring(secondColonSymbol + 4,
					str.Length - secondColonSymbol - 4);

				fileInformation.TimeTotal = ParseTimeSpan(totalTimeText);
			} catch (Exception e) {
				UpdateTextBox("Не удалось выполнить разбор последней строки" +
					Environment.NewLine + "Ошибка: " + e.Message + " " + e.StackTrace, error: true);
			}

			return fileInformation;
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

		private void UpdateTextBox(string message,  bool newSection = false,  bool error = false) {
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

		private List<List<string>> GetCsvFileContent(string filePath) {
			List<List<string>> returnValue = new List<List<string>>();

			if (!IsFileExistAndNotEmpty(filePath)) return returnValue;

			try {
				using (TextFieldParser parser = new TextFieldParser(
					filePath, Encoding.GetEncoding("windows-1251"))) {
					parser.TextFieldType = FieldType.Delimited;
					parser.SetDelimiters(";");
					while (!parser.EndOfData) {
						string[] fields = parser.ReadFields();
						returnValue.Add(fields.ToList());
					}
				}
			} catch (Exception e) {
				UpdateTextBox(e.Message + Environment.NewLine + e.StackTrace);
			}

			return returnValue;
		}

		private bool IsFileExistAndNotEmpty(string filePath) {
			if (File.Exists(filePath)) {
				FileInfo info = new FileInfo(filePath);
				if (info.Length != 0) return true;
			}

			return false;
		}
	}
}
