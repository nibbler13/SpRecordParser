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

		public FileParser(ProgressBar progressBar, TextBox textBox) {
			this.progressBar = progressBar;
			this.textBox = textBox;
			filesInfo = new Dictionary<string, SpRecordFileInformation>();
		}

		public void ParseFiles(List<string> fileNames) {
			UpdateTextBox("Начало анализа");


			foreach (string fileName in fileNames) {
				UpdateTextBox("Файл:" + fileName);
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
			}

			MessageBox.Show("Завершено");
		}

		private void AnalyseFileContentAndAddToDictionary(
			string fileName, List<List<string>> fileContent) {
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
				} else if (line.Count == 9) {

				} else {
					UpdateTextBox("Размер строки " + (row + 1) + " не совпадает с форматом SpRecord");
				}
			}











			for (int row = fileContent.Count - 2; row >= 0; row--) {
				if (row == fileContent.Count - 1) {

				}
			}



			filesInfo.Add(fileName, fileInformation);
		}

		private void ParseLineCreationDate(string text, ref SpRecordFileInformation fileInformation) {
			try {
				int colonSymbol = text.IndexOf(":");
				string creationDate = text.Substring(colonSymbol + 2, text.Length - colonSymbol - 3);
				fileInformation.creationDate = creationDate;
				UpdateTextBox("Дата создания списка: " + creationDate);
			} catch (Exception e) {
				UpdateTextBox("Не удалось выполнить разбор строки с датой создания" + 
					Environment.NewLine + "Ошибка: " + e.Message + " " + e.StackTrace, error: true);
			}
		}

		private void ParseLineWorkstation(string text, ref SpRecordFileInformation fileInformation) {
			try {
				int colonSymbol = text.IndexOf(":");
				string workstationName = text.Substring(colonSymbol + 2, text.Length - colonSymbol - 3);
				fileInformation.workstationName = workstationName;
				UpdateTextBox("Рабочая станция: " + workstationName);
			} catch (Exception e) {
				UpdateTextBox("Не удалось выполнить разбор строки с рабочей станцией" +
					Environment.NewLine + "Ошибка: " + e.Message + " " + e.StackTrace, error: true);
			}
		}

		private void ParseLineAccountingPeriod(string text, ref SpRecordFileInformation fileInformation) {
			try {
				int mark = text.IndexOf("записи");
				string accountingPeriod = text.Substring(mark + 7, text.Length - mark - 7);
				fileInformation.accountingPeriod = accountingPeriod;
				UpdateTextBox("Период отчета: " + accountingPeriod);
			} catch (Exception e) {
				UpdateTextBox("Не удалось выполнить разбор строки с периодом отчета" +
					Environment.NewLine + "Ошибка: " + e.Message + " " + e.StackTrace, error: true);
			}
		}

		private SpRecordFileInformation ParseLastRow(string lastRowText) {
			SpRecordFileInformation fileInformation = new SpRecordFileInformation();

			try {
				int colonSymbol = lastRowText.IndexOf(":");
				int dotSymbol = lastRowText.IndexOf('.');
				string totalRecordsText = lastRowText.Substring(
					colonSymbol + 2, dotSymbol - colonSymbol - 2);

				int totalRecords = 0;
				if (!int.TryParse(totalRecordsText, out totalRecords))
					UpdateTextBox("Не удалось считать общее количество записей", error: false);

				fileInformation.callsTotal = totalRecords;

				int secondColonSymbol = lastRowText.IndexOf("ть:");
				string totalTimeText = lastRowText.Substring(secondColonSymbol + 4,
					lastRowText.Length - secondColonSymbol - 4);

				fileInformation.timeTotal = totalTimeText;
			} catch (Exception e) {
				UpdateTextBox("Не удалось выполнить разбор последней строки" +
					Environment.NewLine + "Ошибка: " + e.Message + " " + e.StackTrace, error: true);
			}

			return fileInformation;
		}

		private void UpdateProgressBar(int percentage) {
			if (progressBar == null) return;
			progressBar.BeginInvoke((MethodInvoker)delegate {
				progressBar.Value = percentage;
			});
		}

		private void UpdateTextBox(
			string message, 
			bool newSection = false, 
			bool error = false) {
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
