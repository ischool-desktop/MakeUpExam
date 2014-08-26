using Aspose.Cells;
using FISCA.Presentation.Controls;
using JHSchool.Data;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MakeUpExam
{
	public partial class MakeUpExamForm : BaseForm
	{
		private int schoolYear, semester;
		private List<string> classIds;

		public MakeUpExamForm(List<string> selectClassIds)
		{
			InitializeComponent();
			// TODO: Complete member initialization
			classIds = selectClassIds;

			schoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
			semester = int.Parse(K12.Data.School.DefaultSemester);

			for (int i = -3; i <= 0; i++)
			{
				cbbSchoolYear.Items.Add(schoolYear + i);
			}
			cbbSemester.Items.Add(1);
			cbbSemester.Items.Add(2);
			cbbSchoolYear.Text = schoolYear + "";
			cbbSemester.Text = semester + "";
		}
		class DataRecord
		{
			public StudentRecord studentRecordData;
			public DomainScore DomainScoreData;
		}
		private void btnPrint_Click(object sender, EventArgs e)
		{
			int schoolYear;
			int semester;
			if (!int.TryParse(cbbSchoolYear.Text, out schoolYear))
			{
				MsgBox.Show("學年度必須選擇為數字");
				return;
			}
			if (!int.TryParse(cbbSemester.Text, out semester))
			{
				MsgBox.Show("學期必須選擇為數字");
				return;
			}

			List<ClassRecord> classList = K12.Data.Class.SelectByIDs(classIds);
			List<string> studentIds = new List<string>();

			string strClassId = string.Join(",", classIds);
			if (strClassId == "")
			{
				MsgBox.Show("無選取班級，請確認是否選取班級");
				return;
			}

			classList.Sort(classsort); //sort class
			Stopwatch sw = Stopwatch.StartNew();
			double total = 0;

			sw.Restart(); //設定計時

			List<StudentRecord> studentAll = K12.Data.Student.SelectByClassIDs(classIds);
			List<JHSemesterScoreRecord> semesterScoreList = JHSemesterScore.SelectBySchoolYearAndSemester(studentAll.Select(x => x.ID).ToList(), schoolYear, semester);
			Dictionary<string, ClassRecord> classIdToRecord = new Dictionary<string,ClassRecord>();
			
			//取得班級名稱
			foreach (ClassRecord cr in classList)
			{
				if (!classIdToRecord.ContainsKey(cr.ID))
				{
					classIdToRecord.Add(cr.ID, cr);
				}
			}

			List<string> domains = new List<string>()
			{
				"國語文", "英語", "數學", "社會", "自然與生活科技", "健康與體育", "藝術與人文", "綜合活動"
			};
			
			Dictionary<string, Dictionary<string,DomainScore>> MakeUpDic = new Dictionary<string,Dictionary<string,DomainScore>>();
			foreach (JHSemesterScoreRecord JHssr in semesterScoreList) 
			{
				foreach (KeyValuePair<string,DomainScore> item in JHssr.Domains)
				{
					if (item.Value.Score.HasValue && item.Value.Score.Value < 60)
					{
						if (!MakeUpDic.ContainsKey(JHssr.RefStudentID))
							MakeUpDic.Add(JHssr.RefStudentID,new Dictionary<string,DomainScore>());
						MakeUpDic[JHssr.RefStudentID].Add(item.Key,item.Value);
						//domains.Add(item.Key);
					}
				}
			}
			//domains = domains.Distinct().ToList();
			//domains.Sort(subSort);

			//各科目的整理
			Dictionary<string, int> subDic = new Dictionary<string,int>();
			int index = 6;
			foreach(string key in domains)
			{
				if(!subDic.ContainsKey(key))
					subDic.Add(key,index);
				index++;
			}

			Workbook wb = new Workbook();
			wb.Open(new MemoryStream(Properties.Resources.Template));
			Worksheet templateSheet = wb.Worksheets["Template"];
			int sheetIndex = 1;
			sheetIndex = wb.Worksheets.AddCopy("Template");

			wb.Worksheets[sheetIndex].Cells[0, 0].PutValue("年級");
			wb.Worksheets[sheetIndex].Cells[0, 1].PutValue("班級");
			wb.Worksheets[sheetIndex].Cells[0, 2].PutValue("座號");
			wb.Worksheets[sheetIndex].Cells[0, 3].PutValue("姓名");
			wb.Worksheets[sheetIndex].Cells[0, 4].PutValue("學年度");
			wb.Worksheets[sheetIndex].Cells[0, 5].PutValue("學期");

			int subIndex = 6;
			foreach (string d in domains)
			{
				wb.Worksheets[sheetIndex].Cells[0, subIndex].PutValue(d);
				subIndex++;
			}

			int rowIndex = 1;
			foreach (StudentRecord sr in studentAll)
			{
				//此學生需要補考
				if (MakeUpDic.ContainsKey(sr.ID))
				{
					if (classIdToRecord.ContainsKey(sr.RefClassID))
					{
						wb.Worksheets[sheetIndex].Cells[rowIndex, 0].PutValue(classIdToRecord[sr.RefClassID].GradeYear);
						wb.Worksheets[sheetIndex].Cells[rowIndex, 1].PutValue(classIdToRecord[sr.RefClassID].Name);
					}
					wb.Worksheets[sheetIndex].Cells[rowIndex, 2].PutValue(sr.SeatNo);
					wb.Worksheets[sheetIndex].Cells[rowIndex, 3].PutValue(sr.Name);
					wb.Worksheets[sheetIndex].Cells[rowIndex, 4].PutValue(cbbSchoolYear.Text);
					wb.Worksheets[sheetIndex].Cells[rowIndex, 5].PutValue(cbbSemester.Text);	
					foreach (string d in domains)
					{
						int columnIndex = subDic[d];
						if (MakeUpDic[sr.ID].ContainsKey(d))
						{
							wb.Worksheets[sheetIndex].Cells[rowIndex, columnIndex].PutValue(MakeUpDic[sr.ID][d].Score);
						}
					}
					rowIndex++;
				}
			}
			total += sw.Elapsed.TotalMilliseconds; //計時標籤 40s
			Console.WriteLine(total);		

			wb.Worksheets.RemoveAt(0);
			SaveFileDialog save = new SaveFileDialog();
			save.Title = "另存新檔";
			save.FileName = cbbSchoolYear.Text + "." + cbbSemester.Text + "補考學生清單.xls";
			save.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
			if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				try
				{
					wb.Save(save.FileName, Aspose.Cells.FileFormatType.Excel2003);
					System.Diagnostics.Process.Start(save.FileName);
				}
				catch
				{
					MessageBox.Show("檔案儲存失敗");
				}
			}
		}

		private int classsort(ClassRecord x, ClassRecord y)
		{
			string xx = (x.GradeYear + "").PadLeft(3, '0');
			xx += x.Name.PadLeft(20, '0');

			string yy = (y.GradeYear + "").PadLeft(3, '0');
			yy += y.Name.PadLeft(20, '0');

			return xx.CompareTo(yy);
		}

		private void btnQuit_Click(object sender, EventArgs e)
		{
			this.Close();
		}
	}
}
