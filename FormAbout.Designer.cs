namespace SpRecordParser {
	partial class FormAbout {
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing) {
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent() {
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormAbout));
			this.labelBottomHelp = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.pictureBox3 = new System.Windows.Forms.PictureBox();
			this.pictureBox2 = new System.Windows.Forms.PictureBox();
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
			this.SuspendLayout();
			// 
			// labelBottomHelp
			// 
			this.labelBottomHelp.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.labelBottomHelp.Location = new System.Drawing.Point(12, 9);
			this.labelBottomHelp.Name = "labelBottomHelp";
			this.labelBottomHelp.Size = new System.Drawing.Size(390, 94);
			this.labelBottomHelp.TabIndex = 8;
			this.labelBottomHelp.Text = resources.GetString("labelBottomHelp.Text");
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(12, 298);
			this.label1.Margin = new System.Windows.Forms.Padding(3, 10, 3, 0);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(390, 46);
			this.label1.TabIndex = 11;
			this.label1.Text = "Порядок столбцов: \r\nНазвание канала, Время записи, Длительность, Тип записи, Номе" +
    "р абонента, Ext, CO, Тип CO, Комментарий:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(12, 429);
			this.label2.Margin = new System.Windows.Forms.Padding(3, 10, 3, 0);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(218, 13);
			this.label2.TabIndex = 13;
			this.label2.Text = "Требуется версия SpRecord не ниже 3.99";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(12, 469);
			this.label3.Margin = new System.Windows.Forms.Padding(3, 10, 3, 0);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(364, 13);
			this.label3.TabIndex = 14;
			this.label3.Text = "Автор: Грашкин Павел / nn-admin@bzklinika.ru / отдел бизнес-анализа";
			// 
			// pictureBox3
			// 
			this.pictureBox3.Image = global::SpRecordParser.Properties.Resources.Save4;
			this.pictureBox3.Location = new System.Drawing.Point(12, 347);
			this.pictureBox3.Name = "pictureBox3";
			this.pictureBox3.Size = new System.Drawing.Size(390, 69);
			this.pictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
			this.pictureBox3.TabIndex = 12;
			this.pictureBox3.TabStop = false;
			// 
			// pictureBox2
			// 
			this.pictureBox2.Image = global::SpRecordParser.Properties.Resources.Save2;
			this.pictureBox2.Location = new System.Drawing.Point(12, 210);
			this.pictureBox2.Name = "pictureBox2";
			this.pictureBox2.Size = new System.Drawing.Size(390, 75);
			this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
			this.pictureBox2.TabIndex = 10;
			this.pictureBox2.TabStop = false;
			// 
			// pictureBox1
			// 
			this.pictureBox1.Image = global::SpRecordParser.Properties.Resources.Save1;
			this.pictureBox1.Location = new System.Drawing.Point(85, 106);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(244, 98);
			this.pictureBox1.TabIndex = 9;
			this.pictureBox1.TabStop = false;
			// 
			// FormAbout
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(414, 491);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.pictureBox3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.pictureBox2);
			this.Controls.Add(this.pictureBox1);
			this.Controls.Add(this.labelBottomHelp);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimumSize = new System.Drawing.Size(430, 530);
			this.Name = "FormAbout";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "О программе";
			((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label labelBottomHelp;
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.PictureBox pictureBox2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.PictureBox pictureBox3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
	}
}