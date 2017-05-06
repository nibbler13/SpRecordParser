using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;

namespace SpRecordParser {
	public partial class FormMain : Form {
		Thread parsingThread;
		List<string> selectedFiles;

		public FormMain() {
			InitializeComponent();

			listViewFiles.ItemSelectionChanged += ListViewFiles_ItemSelectionChanged;
			this.FormClosing += Form1_FormClosing;
		}

		private void Form1_FormClosing(object sender, FormClosingEventArgs e) {
			if (parsingThread != null && parsingThread.IsAlive)
				e.Cancel = true;
		}

		private void ListViewFiles_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e) {
			buttonDelete.Enabled = listViewFiles.SelectedItems.Count > 0;
		}

		private void buttonAdd_Click(object sender, EventArgs e) {
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.Filter = "Список записей SpRecord (*.csv)|*.csv";
			openFileDialog.CheckFileExists = true;
			openFileDialog.CheckPathExists = true;
			openFileDialog.Multiselect = true;
			openFileDialog.RestoreDirectory = true;

			if (openFileDialog.ShowDialog() == DialogResult.OK) {
				foreach (string fileName in openFileDialog.FileNames) {
					if (!fileName.EndsWith(".csv"))
						continue;

					ListViewItem item = new ListViewItem(fileName);
					listViewFiles.Items.Add(item);
				}

				if (listViewFiles.Items.Count > 0)
					buttonAnalyse.Enabled = true;
			}

		}

		private void buttonDelete_Click(object sender, EventArgs e) {
			foreach(ListViewItem item in listViewFiles.SelectedItems) {
				listViewFiles.Items.Remove(item);
			}

			if (listViewFiles.Items.Count == 0)
				buttonAnalyse.Enabled = false;
		}

		private void buttonAnalyse_Click(object sender, EventArgs e) {
			buttonAnalyse.Visible = false;
			buttonAdd.Visible = false;
			buttonDelete.Visible = false;
			labelListTitle.Visible = false;
			labelBottomHelp.Visible = false;
			listViewFiles.Visible = false;
			
			textBox.Visible = true;
			progressBar.Visible = true;

			selectedFiles = new List<string>();
			foreach (ListViewItem item in listViewFiles.Items) {
				selectedFiles.Add(item.Text);
			}

			parsingThread = new Thread(StartParsing);
			parsingThread.Start();
		}

		public void StartParsing() {
			FileParser fileParser = new FileParser(progressBar, textBox);
			fileParser.ParseFiles(selectedFiles);
		}
	}
}
