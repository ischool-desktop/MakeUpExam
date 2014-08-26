namespace MakeUpExam
{
	partial class MakeUpExamForm
	{
		/// <summary>
		/// 設計工具所需的變數。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// 清除任何使用中的資源。
		/// </summary>
		/// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form 設計工具產生的程式碼

		/// <summary>
		/// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
		/// 修改這個方法的內容。
		/// </summary>
		private void InitializeComponent()
		{
			this.btnQuit = new DevComponents.DotNetBar.ButtonX();
			this.labelX1 = new DevComponents.DotNetBar.LabelX();
			this.cbbSchoolYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
			this.cbbSemester = new DevComponents.DotNetBar.Controls.ComboBoxEx();
			this.btnPrint = new DevComponents.DotNetBar.ButtonX();
			this.labelX2 = new DevComponents.DotNetBar.LabelX();
			this.labelX3 = new DevComponents.DotNetBar.LabelX();
			this.SuspendLayout();
			// 
			// btnQuit
			// 
			this.btnQuit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
			this.btnQuit.BackColor = System.Drawing.Color.Transparent;
			this.btnQuit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
			this.btnQuit.Location = new System.Drawing.Point(202, 74);
			this.btnQuit.Name = "btnQuit";
			this.btnQuit.Size = new System.Drawing.Size(75, 23);
			this.btnQuit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
			this.btnQuit.TabIndex = 1;
			this.btnQuit.Text = "取消";
			this.btnQuit.Click += new System.EventHandler(this.btnQuit_Click);
			// 
			// labelX1
			// 
			this.labelX1.BackColor = System.Drawing.Color.Transparent;
			// 
			// 
			// 
			this.labelX1.BackgroundStyle.Class = "";
			this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
			this.labelX1.Location = new System.Drawing.Point(12, 12);
			this.labelX1.Name = "labelX1";
			this.labelX1.Size = new System.Drawing.Size(115, 23);
			this.labelX1.TabIndex = 2;
			this.labelX1.Text = "列印補考學生清單";
			// 
			// cbbSchoolYear
			// 
			this.cbbSchoolYear.DisplayMember = "Text";
			this.cbbSchoolYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
			this.cbbSchoolYear.FormattingEnabled = true;
			this.cbbSchoolYear.ItemHeight = 19;
			this.cbbSchoolYear.Location = new System.Drawing.Point(69, 41);
			this.cbbSchoolYear.Name = "cbbSchoolYear";
			this.cbbSchoolYear.Size = new System.Drawing.Size(88, 25);
			this.cbbSchoolYear.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
			this.cbbSchoolYear.TabIndex = 3;
			// 
			// cbbSemester
			// 
			this.cbbSemester.DisplayMember = "Text";
			this.cbbSemester.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
			this.cbbSemester.FormattingEnabled = true;
			this.cbbSemester.ItemHeight = 19;
			this.cbbSemester.Location = new System.Drawing.Point(202, 41);
			this.cbbSemester.Name = "cbbSemester";
			this.cbbSemester.Size = new System.Drawing.Size(75, 25);
			this.cbbSemester.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
			this.cbbSemester.TabIndex = 4;
			// 
			// btnPrint
			// 
			this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
			this.btnPrint.BackColor = System.Drawing.Color.Transparent;
			this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
			this.btnPrint.Location = new System.Drawing.Point(121, 74);
			this.btnPrint.Name = "btnPrint";
			this.btnPrint.Size = new System.Drawing.Size(75, 23);
			this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
			this.btnPrint.TabIndex = 0;
			this.btnPrint.Text = "確認";
			this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
			// 
			// labelX2
			// 
			this.labelX2.BackColor = System.Drawing.Color.Transparent;
			// 
			// 
			// 
			this.labelX2.BackgroundStyle.Class = "";
			this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
			this.labelX2.Location = new System.Drawing.Point(12, 41);
			this.labelX2.Name = "labelX2";
			this.labelX2.Size = new System.Drawing.Size(51, 23);
			this.labelX2.TabIndex = 5;
			this.labelX2.Text = "學年度";
			// 
			// labelX3
			// 
			this.labelX3.BackColor = System.Drawing.Color.Transparent;
			// 
			// 
			// 
			this.labelX3.BackgroundStyle.Class = "";
			this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
			this.labelX3.Location = new System.Drawing.Point(163, 41);
			this.labelX3.Name = "labelX3";
			this.labelX3.Size = new System.Drawing.Size(33, 23);
			this.labelX3.TabIndex = 6;
			this.labelX3.Text = "學期";
			// 
			// MakeUpExamForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(288, 105);
			this.Controls.Add(this.labelX3);
			this.Controls.Add(this.labelX2);
			this.Controls.Add(this.cbbSemester);
			this.Controls.Add(this.cbbSchoolYear);
			this.Controls.Add(this.labelX1);
			this.Controls.Add(this.btnQuit);
			this.Controls.Add(this.btnPrint);
			this.DoubleBuffered = true;
			this.Name = "MakeUpExamForm";
			this.Text = "補考學生清單";
			this.ResumeLayout(false);

		}

		#endregion

		private DevComponents.DotNetBar.ButtonX btnQuit;
		private DevComponents.DotNetBar.LabelX labelX1;
		private DevComponents.DotNetBar.Controls.ComboBoxEx cbbSchoolYear;
		private DevComponents.DotNetBar.Controls.ComboBoxEx cbbSemester;
		private DevComponents.DotNetBar.ButtonX btnPrint;
		private DevComponents.DotNetBar.LabelX labelX2;
		private DevComponents.DotNetBar.LabelX labelX3;
	}
}

