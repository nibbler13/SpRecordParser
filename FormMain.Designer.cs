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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
			this.toolStripMenuItemSettings = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripMenuItemAbout = new System.Windows.Forms.ToolStripMenuItem();
			this.buttonDelete = new System.Windows.Forms.Button();
			this.listViewFiles = new System.Windows.Forms.ListView();
			this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
			this.buttonAdd = new System.Windows.Forms.Button();
			this.labelListTitle = new System.Windows.Forms.Label();
			this.buttonAnalyse = new System.Windows.Forms.Button();
			this.progressBar = new System.Windows.Forms.ProgressBar();
			this.textBox = new System.Windows.Forms.TextBox();
			this.настройкиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.оПрограммеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.menuStrip1 = new System.Windows.Forms.MenuStrip();
			this.dfdsfToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.оПрограммеToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
			this.menuStrip1.SuspendLayout();
			this.SuspendLayout();
			// 
			// toolStripMenuItemSettings
			// 
			this.toolStripMenuItemSettings.Name = "toolStripMenuItemSettings";
			this.toolStripMenuItemSettings.Size = new System.Drawing.Size(79, 20);
			this.toolStripMenuItemSettings.Text = "Настройки";
			this.toolStripMenuItemSettings.Click += new System.EventHandler(this.settingsToolStripMenuItem_Click);
			// 
			// toolStripMenuItemAbout
			// 
			this.toolStripMenuItemAbout.Name = "toolStripMenuItemAbout";
			this.toolStripMenuItemAbout.Size = new System.Drawing.Size(94, 20);
			this.toolStripMenuItemAbout.Text = "О программе";
			this.toolStripMenuItemAbout.Click += new System.EventHandler(this.aboutToolStripMenuItem1_Click);
			// 
			// buttonDelete
			// 
			this.buttonDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.buttonDelete.Enabled = false;
			this.buttonDelete.Location = new System.Drawing.Point(93, 456);
			this.buttonDelete.Name = "buttonDelete";
			this.buttonDelete.Size = new System.Drawing.Size(75, 23);
			this.buttonDelete.TabIndex = 2;
			this.buttonDelete.Text = "Удалить";
			this.buttonDelete.UseVisualStyleBackColor = true;
			this.buttonDelete.Click += new System.EventHandler(this.buttonDelete_Click);
			// 
			// listViewFiles
			// 
			this.listViewFiles.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.listViewFiles.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
			this.listViewFiles.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.listViewFiles.Location = new System.Drawing.Point(12, 27);
			this.listViewFiles.Name = "listViewFiles";
			this.listViewFiles.Size = new System.Drawing.Size(390, 400);
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
			this.buttonAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.buttonAdd.Location = new System.Drawing.Point(12, 456);
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
			this.labelListTitle.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
			this.labelListTitle.Location = new System.Drawing.Point(132, 430);
			this.labelListTitle.Margin = new System.Windows.Forms.Padding(3, 0, 3, 10);
			this.labelListTitle.Name = "labelListTitle";
			this.labelListTitle.Size = new System.Drawing.Size(151, 13);
			this.labelListTitle.TabIndex = 4;
			this.labelListTitle.Text = "Список файлов для анализа";
			// 
			// buttonAnalyse
			// 
			this.buttonAnalyse.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
			this.buttonAnalyse.Enabled = false;
			this.buttonAnalyse.Location = new System.Drawing.Point(267, 456);
			this.buttonAnalyse.Name = "buttonAnalyse";
			this.buttonAnalyse.Size = new System.Drawing.Size(135, 23);
			this.buttonAnalyse.TabIndex = 3;
			this.buttonAnalyse.Text = "Выполнить анализ";
			this.buttonAnalyse.UseVisualStyleBackColor = true;
			this.buttonAnalyse.Click += new System.EventHandler(this.buttonAnalyse_Click);
			// 
			// progressBar
			// 
			this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.progressBar.Location = new System.Drawing.Point(12, 456);
			this.progressBar.Name = "progressBar";
			this.progressBar.Size = new System.Drawing.Size(390, 23);
			this.progressBar.TabIndex = 8;
			this.progressBar.Visible = false;
			// 
			// textBox
			// 
			this.textBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.textBox.Location = new System.Drawing.Point(12, 27);
			this.textBox.Multiline = true;
			this.textBox.Name = "textBox";
			this.textBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.textBox.Size = new System.Drawing.Size(390, 423);
			this.textBox.TabIndex = 9;
			this.textBox.Visible = false;
			// 
			// настройкиToolStripMenuItem
			// 
			this.настройкиToolStripMenuItem.Name = "настройкиToolStripMenuItem";
			this.настройкиToolStripMenuItem.Size = new System.Drawing.Size(79, 20);
			this.настройкиToolStripMenuItem.Text = "Настройки";
			this.настройкиToolStripMenuItem.Click += new System.EventHandler(this.settingsToolStripMenuItem_Click);
			// 
			// оПрограммеToolStripMenuItem
			// 
			this.оПрограммеToolStripMenuItem.Name = "оПрограммеToolStripMenuItem";
			this.оПрограммеToolStripMenuItem.Size = new System.Drawing.Size(94, 20);
			this.оПрограммеToolStripMenuItem.Text = "О программе";
			this.оПрограммеToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem1_Click);
			// 
			// menuStrip1
			// 
			this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.dfdsfToolStripMenuItem,
            this.оПрограммеToolStripMenuItem1});
			this.menuStrip1.Location = new System.Drawing.Point(0, 0);
			this.menuStrip1.Name = "menuStrip1";
			this.menuStrip1.Size = new System.Drawing.Size(414, 24);
			this.menuStrip1.TabIndex = 10;
			this.menuStrip1.Text = "menuStrip1";
			// 
			// dfdsfToolStripMenuItem
			// 
			this.dfdsfToolStripMenuItem.Name = "dfdsfToolStripMenuItem";
			this.dfdsfToolStripMenuItem.Size = new System.Drawing.Size(79, 20);
			this.dfdsfToolStripMenuItem.Text = "Настройки";
			this.dfdsfToolStripMenuItem.Click += new System.EventHandler(this.settingsToolStripMenuItem_Click);
			// 
			// оПрограммеToolStripMenuItem1
			// 
			this.оПрограммеToolStripMenuItem1.Name = "оПрограммеToolStripMenuItem1";
			this.оПрограммеToolStripMenuItem1.Size = new System.Drawing.Size(94, 20);
			this.оПрограммеToolStripMenuItem1.Text = "О программе";
			this.оПрограммеToolStripMenuItem1.Click += new System.EventHandler(this.aboutToolStripMenuItem1_Click);
			// 
			// FormMain
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(414, 491);
			this.Controls.Add(this.buttonAnalyse);
			this.Controls.Add(this.labelListTitle);
			this.Controls.Add(this.buttonAdd);
			this.Controls.Add(this.listViewFiles);
			this.Controls.Add(this.buttonDelete);
			this.Controls.Add(this.progressBar);
			this.Controls.Add(this.textBox);
			this.Controls.Add(this.menuStrip1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimumSize = new System.Drawing.Size(430, 530);
			this.Name = "FormMain";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Анализ журналов звонков SpRecord";
			this.menuStrip1.ResumeLayout(false);
			this.menuStrip1.PerformLayout();
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
		private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemSettings;
		private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemAbout;
		private System.Windows.Forms.ToolStripMenuItem настройкиToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem оПрограммеToolStripMenuItem;
		private System.Windows.Forms.MenuStrip menuStrip1;
		private System.Windows.Forms.ToolStripMenuItem dfdsfToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem оПрограммеToolStripMenuItem1;
	}
}

