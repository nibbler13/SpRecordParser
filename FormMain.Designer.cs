namespace SpRecordParser {
	partial class FormMain {
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		/// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
		protected override void Dispose(bool disposing) {
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Код, автоматически созданный конструктором форм Windows

		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		private void InitializeComponent() {
			this.buttonDelete = new System.Windows.Forms.Button();
			this.listViewFiles = new System.Windows.Forms.ListView();
			this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
			this.buttonAdd = new System.Windows.Forms.Button();
			this.labelListTitle = new System.Windows.Forms.Label();
			this.labelBottomHelp = new System.Windows.Forms.Label();
			this.buttonAnalyse = new System.Windows.Forms.Button();
			this.progressBar = new System.Windows.Forms.ProgressBar();
			this.textBox = new System.Windows.Forms.TextBox();
			this.SuspendLayout();
			// 
			// buttonDelete
			// 
			this.buttonDelete.Enabled = false;
			this.buttonDelete.Location = new System.Drawing.Point(16, 423);
			this.buttonDelete.Name = "buttonDelete";
			this.buttonDelete.Size = new System.Drawing.Size(75, 23);
			this.buttonDelete.TabIndex = 2;
			this.buttonDelete.Text = "Удалить";
			this.buttonDelete.UseVisualStyleBackColor = true;
			this.buttonDelete.Click += new System.EventHandler(this.buttonDelete_Click);
			// 
			// listViewFiles
			// 
			this.listViewFiles.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
			this.listViewFiles.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.listViewFiles.Location = new System.Drawing.Point(16, 29);
			this.listViewFiles.Name = "listViewFiles";
			this.listViewFiles.Size = new System.Drawing.Size(385, 315);
			this.listViewFiles.TabIndex = 4;
			this.listViewFiles.UseCompatibleStateImageBehavior = false;
			this.listViewFiles.View = System.Windows.Forms.View.Details;
			// 
			// columnHeader1
			// 
			this.columnHeader1.Text = "Имя файла";
			this.columnHeader1.Width = 379;
			// 
			// buttonAdd
			// 
			this.buttonAdd.Location = new System.Drawing.Point(326, 423);
			this.buttonAdd.Name = "buttonAdd";
			this.buttonAdd.Size = new System.Drawing.Size(75, 23);
			this.buttonAdd.TabIndex = 1;
			this.buttonAdd.Text = "Добавить";
			this.buttonAdd.UseVisualStyleBackColor = true;
			this.buttonAdd.Click += new System.EventHandler(this.buttonAdd_Click);
			// 
			// labelListTitle
			// 
			this.labelListTitle.AutoSize = true;
			this.labelListTitle.Location = new System.Drawing.Point(13, 13);
			this.labelListTitle.Name = "labelListTitle";
			this.labelListTitle.Size = new System.Drawing.Size(154, 13);
			this.labelListTitle.TabIndex = 4;
			this.labelListTitle.Text = "Список файлов для анализа:";
			// 
			// labelBottomHelp
			// 
			this.labelBottomHelp.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
			this.labelBottomHelp.Location = new System.Drawing.Point(13, 347);
			this.labelBottomHelp.Name = "labelBottomHelp";
			this.labelBottomHelp.Size = new System.Drawing.Size(388, 73);
			this.labelBottomHelp.TabIndex = 7;
			this.labelBottomHelp.Text = "Требуются файлы формата CSV, созданные в программе SpRecord.\r\n\r\nПорядок столбцов:" +
    " \r\nНазвание канала, Время записи, Длительность, Тип записи, Номер абонента, Ext," +
    " CO, Тип CO, Комментарий";
			// 
			// buttonAnalyse
			// 
			this.buttonAnalyse.Enabled = false;
			this.buttonAnalyse.Location = new System.Drawing.Point(139, 453);
			this.buttonAnalyse.Name = "buttonAnalyse";
			this.buttonAnalyse.Size = new System.Drawing.Size(135, 23);
			this.buttonAnalyse.TabIndex = 3;
			this.buttonAnalyse.Text = "Выполнить анализ";
			this.buttonAnalyse.UseVisualStyleBackColor = true;
			this.buttonAnalyse.Click += new System.EventHandler(this.buttonAnalyse_Click);
			// 
			// progressBar
			// 
			this.progressBar.Location = new System.Drawing.Point(14, 453);
			this.progressBar.Name = "progressBar";
			this.progressBar.Size = new System.Drawing.Size(385, 23);
			this.progressBar.TabIndex = 8;
			this.progressBar.Visible = false;
			// 
			// textBox
			// 
			this.textBox.Location = new System.Drawing.Point(12, 12);
			this.textBox.Multiline = true;
			this.textBox.Name = "textBox";
			this.textBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.textBox.Size = new System.Drawing.Size(389, 434);
			this.textBox.TabIndex = 9;
			this.textBox.Visible = false;
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(413, 488);
			this.Controls.Add(this.buttonAnalyse);
			this.Controls.Add(this.labelBottomHelp);
			this.Controls.Add(this.labelListTitle);
			this.Controls.Add(this.buttonAdd);
			this.Controls.Add(this.listViewFiles);
			this.Controls.Add(this.buttonDelete);
			this.Controls.Add(this.progressBar);
			this.Controls.Add(this.textBox);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "Form1";
			this.Text = "Анализ журналов звонков SpRecord";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button buttonDelete;
		private System.Windows.Forms.ListView listViewFiles;
		private System.Windows.Forms.ColumnHeader columnHeader1;
		private System.Windows.Forms.Button buttonAdd;
		private System.Windows.Forms.Button buttonAnalyse;
		private System.Windows.Forms.ProgressBar progressBar;
		private System.Windows.Forms.TextBox textBox;
		private System.Windows.Forms.Label labelListTitle;
		private System.Windows.Forms.Label labelBottomHelp;
	}
}

