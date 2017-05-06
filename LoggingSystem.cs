﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpRecordParser {
	class LoggingSystem {
		private const string LOG_FILE_NAME = "SpRecordParser*.log";
		private const int MAX_LOGFILES_QUANTITY = 5;

		public static void LogMessageToFile(string msg) {
			string today = DateTime.Now.ToString("yyyyMMdd");
			string logFileName = Directory.GetCurrentDirectory() + "\\" + LOG_FILE_NAME.Replace("*", today);
			try {
				using (System.IO.StreamWriter sw = System.IO.File.AppendText(logFileName)) {
					string logLine = System.String.Format("{0:G}: {1}", System.DateTime.Now, msg);
					sw.WriteLine(logLine);
				}
			} catch (Exception e) {
				Console.WriteLine("Cannot write to log file: " + logFileName + " " + e.StackTrace + " " + e.Message);
			}

			Console.WriteLine(msg);
			CheckAndCleanOldFiles();
		}

		private static void CheckAndCleanOldFiles() {
			try {
				DirectoryInfo dirInfo = new DirectoryInfo(Directory.GetCurrentDirectory());
				FileInfo[] files = dirInfo.GetFiles(LOG_FILE_NAME).OrderBy(p => p.CreationTime).ToArray();
				if (files.Length > MAX_LOGFILES_QUANTITY) {
					for (int i = 0; i < files.Length - MAX_LOGFILES_QUANTITY; i++) {
						files[i].Delete();
					}
				}
			} catch (Exception e) {
				Console.WriteLine("Cannot delete old lig files: " + e.StackTrace + " " + e.Message);
			}
		}
	}
}
