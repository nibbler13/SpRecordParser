﻿using Microsoft.VisualBasic.FileIO;
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

			int oneFileProgress = 90 / fileNames.Count;
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
				MessageBox.Show("Результирующий файл пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

			UpdateTextBox("Выгрузка данных в Excel", newSection: true);
			ExcelWriter.WriteToExcel(filesInfo);

			UpdateProgressBar(100);

			MessageBox.Show("Анализ завершен", "SpRecordParser", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
				if (line.Count == 1) {
					string text = line[0];

					if (text.StartsWith("Выбраны записи"))
						ParseLineAccountingPeriod(text, ref fileInformation);

					if (text.StartsWith("Рабочая станция"))
						ParseLineWorkstation(text, ref fileInformation);

					if (text.StartsWith("Отчет создан"))
						ParseLineCreationDate(text, ref fileInformation);
				} else if (line.Count >= 9) {
					if (line[0].Equals("Название канала"))
						continue;

					TimeSpan duration = ParseTimeSpan(line[2]);
					string type = line[3];

					switch (type) {
						case "Принятый":
							fileInformation.callsAccepted++;
							fileInformation.timeAccepted = fileInformation.timeAccepted.Add(duration);
							break;
						case "Набранный":
							fileInformation.callsDialed++;
							fileInformation.timeDialed = fileInformation.timeDialed.Add(duration);
							break;
						case "Непринятый":
							if (duration.TotalSeconds <= 5) {
								fileInformation.callsAccidential++;
								fileInformation.timeAccidential = fileInformation.timeAccidential.Add(duration);
								fileContent[row].Add("Accidential");
							} else {
								fileInformation.callsMissed++;
								fileInformation.timeMissed = fileInformation.timeMissed.Add(duration);
								AnalyseMissedCall(row, ref fileInformation, ref fileContent);
							}
							break;
						default:
							UpdateTextBox("Неизвестный тип звонка: " + type);
							break;
					}
				} else {
					UpdateTextBox("Размер строки " + (row + 1) + " не совпадает с форматом SpRecord");
				}
			}

			fileInformation.fileContent = fileContent;
			filesInfo.Add(fileName, fileInformation);
		}

		private void AnalyseMissedCall(
			int row, ref SpRecordFileInformation fileInformation, ref List<List<string>> fileContent) {
			//UpdateTextBox("Анализ пропущенного звонка, строка: " + (row + 1) + Environment.NewLine +
			//	string.Join(";", fileContent[row]));

			DateTime missedTime;
			if (!DateTime.TryParse(fileContent[row][1], out missedTime)) {
				fileContent[row].Add("Invalid date/time");
				UpdateTextBox("Не удалось разобрать время звонка, строка: " + (row + 1) +
					" значение: " + fileContent[row][1]);
				return;
			}

			string[] phoneNumbers = SplitPhoneNumbers(fileContent[row][4]);
			string callerNumber = phoneNumbers[0];

			if (string.IsNullOrEmpty(callerNumber)) {
				fileContent[row].Add("Wrong caller phone number");
				UpdateTextBox("Номер звонившего не удалось определить");
				return;
			}

			//надо проверять номера телефонов без кода города
			//мобильные без 8 спереди
			try {
				if (callerNumber.StartsWith("89")) {
					callerNumber = callerNumber.Substring(1);
				} else {
					callerNumber = callerNumber.Substring(callerNumber.Length - 7);
				}
			} catch (Exception e) {
				LoggingSystem.LogMessageToFile(e.Message);
				LoggingSystem.LogMessageToFile(e.StackTrace);
			}

			int callBackTries = 0;
			bool regulationsObserved = true;
			bool registryCallBackSucceded = false;
			bool conversationTookPlace = false;
			DateTime lastCallTime = missedTime;

			for (int i = row - 1; i >= 0; i--) {
				if (fileContent[i].Count < 9)
					break;

				if (fileContent[i][0].Equals("Название канала"))
					break;

				DateTime callDate;
				if (!DateTime.TryParse(fileContent[i][1], out callDate)) {
					UpdateTextBox("Не удалось разобрать время звонка, строка: " + (i + 1) + 
						" значение: " + fileContent[i][1]);
					continue;
				}
				
				if (!missedTime.Date.Equals(callDate.Date))
					break;

				string callPhoneNumbers = fileContent[i][4];
				if (!callPhoneNumbers.Contains(callerNumber))
					continue;
				
				lastCallTime = callDate;

				string callType = fileContent[i][3];

				if (callType.Equals("Непринятый"))
					break;
				
				fileContent[i].Add("Связка с пропущенным звонком");
				fileContent[i].Add("Строка: " + (row + 1));

				if (callType.Equals("Принятый")) {
					conversationTookPlace = true;
					break;
				} else if (callType.Equals("Набранный")) {
					string comment = fileContent[row][8];
					if (comment.Equals("Вызываемый абонент не ответил.")) {
						callBackTries++;

						double minutesAfterMissedCall = callDate.Subtract(missedTime).TotalMinutes;
						if (callBackTries == 1 && minutesAfterMissedCall > 5.0 ||
							callBackTries == 2 && minutesAfterMissedCall > 20.0 ||
							callBackTries == 3 && minutesAfterMissedCall > 35.0)
							regulationsObserved = false;
					} else {
						conversationTookPlace = true;
						registryCallBackSucceded = true;
						break;
					}
				}
			}

			string result = "";

			if (!conversationTookPlace) {
				result = "Разговор с пациентом не состоялся, ";
				fileInformation.callsBackNot++;

				if (callBackTries == 0) {
					regulationsObserved = false;
					result += "пациенту не пытались перезвонить";
				} else if (callBackTries < 3) {
					regulationsObserved = false;
					result += "пациенту пытались перезвонить менее 3 раз";
				} else {
					result += "пациенту пытались перезвонить 3 или более раз";
				}
			} else {
				double minutesAfterMissedCall = lastCallTime.Subtract(missedTime).TotalMinutes;
				if (callBackTries == 0 && minutesAfterMissedCall > 5.0 ||
					callBackTries == 1 && minutesAfterMissedCall > 20.0 ||
					callBackTries == 2 && minutesAfterMissedCall > 35.0)
					regulationsObserved = false;

				if (registryCallBackSucceded) {
					result = "Регистратура перезвонила пациенту";
					fileInformation.callsBackByRegistry++;
				} else {
					result = "Пациент перезвонил самостоятельно";
					fileInformation.callsBackByPatient++;
				}
			}

			if (regulationsObserved) {
				fileInformation.missedCallsRegulationObserved++;
				result += ", регламент соблюден";
			} else {
				fileInformation.missedCallsRegulationNotObserved++;
				result += ", регламент нарушен";
			}

			fileContent[row].Add(result);
		}

		private string[] SplitPhoneNumbers(string str) {
			string[] phoneNumbers = new string[2];

			if (!str.Contains(" -> "))
				return phoneNumbers;

			try {
				phoneNumbers = str.Split(new[] { " -> " }, StringSplitOptions.None);
			} catch (Exception e) {
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
				fileInformation.creationDate = creationDate;
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
				fileInformation.workstationName = workstationName;
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
				fileInformation.accountingPeriod = accountingPeriod;
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

				fileInformation.callsTotal = totalRecords;

				int secondColonSymbol = str.IndexOf("ть:");
				string totalTimeText = str.Substring(secondColonSymbol + 4,
					str.Length - secondColonSymbol - 4);

				fileInformation.timeTotal = ParseTimeSpan(totalTimeText);
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
