using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;

namespace SpRecordParser {
	public partial class FormMain : Form {
		Thread parsingThread;
		private List<Control> controlsToEnable;
		private List<Control> controlsToVisible;

		public FormMain() {
			InitializeComponent();

			listViewFiles.ItemSelectionChanged += ListViewFiles_ItemSelectionChanged;
			this.FormClosing += Form1_FormClosing;

			controlsToEnable = new List<Control>() {
				buttonAnalyse
			};

			controlsToVisible = new List<Control>() {
				buttonAnalyse,
				buttonAdd,
				buttonDelete,
				labelListTitle,
				listViewFiles
			};
		}

		private void Form1_FormClosing(object sender, FormClosingEventArgs e) {
			if (parsingThread != null && parsingThread.IsAlive)
				e.Cancel = true;
		}

		private void ListViewFiles_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e) {
			buttonDelete.Enabled = listViewFiles.SelectedItems.Count > 0;
		}

		private void ButtonAdd_Click(object sender, EventArgs e) {
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
					foreach (Control control in controlsToEnable)
						control.Enabled = true;
			}

		}

		private void ButtonDelete_Click(object sender, EventArgs e) {
			foreach(ListViewItem item in listViewFiles.SelectedItems)
				listViewFiles.Items.Remove(item);

			if (listViewFiles.Items.Count == 0)
				foreach (Control control in controlsToEnable)
					control.Enabled = false;
		}

		private void ButtonAnalyse_Click(object sender, EventArgs e) {
			foreach (Control control in controlsToVisible)
				control.Visible = false;
			
			textBox.Visible = true;
			progressBar.Visible = true;

			List<string> selectedFiles = new List<string>();
			foreach (ListViewItem item in listViewFiles.Items)
				selectedFiles.Add(item.Text);

			parsingThread = new Thread(() => StartParsing(selectedFiles));
			parsingThread.Start();
		}

		public void StartParsing(List<string> selectedFiles) {
			FileParser fileParser = new FileParser(progressBar, textBox);
			fileParser.ParseFiles(selectedFiles);
		}

		private void SettingsToolStripMenuItem_Click(object sender, EventArgs e) {
		FormSettings formSettings = new FormSettings();
		formSettings.ShowDialog();
	}

		private void AboutToolStripMenuItem1_Click(object sender, EventArgs e) {
			FormAbout formAbout = new FormAbout();
			formAbout.ShowDialog();
		}
	}
}
