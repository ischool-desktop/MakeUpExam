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
		private int _schoolYear, _semester;
		private List<string> _classIds;

        private BackgroundWorker _BW;

        List<string> _domainNameSort = new List<string>()
			{
				"國語文", "英語", "數學", "社會", "自然與生活科技", "健康與體育", "藝術與人文", "綜合活動"
			};

        List<string> _subjNameSort = new List<string>()
			{
				"國語文", "國文", "英文","英語","數學","理化","地理","歷史","生物"
			};

		public MakeUpExamForm(List<string> selectClassIds)
		{
			InitializeComponent();
			// TODO: Complete member initialization
			_classIds = selectClassIds;

			_schoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
			_semester = int.Parse(K12.Data.School.DefaultSemester);

			for (int i = -3; i <= 0; i++)
			{
				cbbSchoolYear.Items.Add(_schoolYear + i);
			}
			cbbSemester.Items.Add(1);
			cbbSemester.Items.Add(2);
			cbbSchoolYear.Text = _schoolYear + "";
			cbbSemester.Text = _semester + "";

            _BW = new BackgroundWorker();

            _BW.DoWork += new DoWorkEventHandler(Dowork);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Completed);
		}

        private void Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            btnPrint.Enabled = true;

            Workbook wb = e.Result as Workbook;
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
                this.Close();
            }
        }

        private void Dowork(object sender, DoWorkEventArgs e)
        {
            List<ClassRecord> classRecords = K12.Data.Class.SelectByIDs(_classIds);

            //Stopwatch sw = Stopwatch.StartNew();
            //double total = 0;

            List<StudentRecord> studentRecords = K12.Data.Student.SelectByClassIDs(_classIds);
            List<JHSemesterScoreRecord> semesterScoreList = JHSemesterScore.SelectBySchoolYearAndSemester(studentRecords.Select(x => x.ID).ToList(), _schoolYear, _semester);
            
            Dictionary<string, ClassRecord> classIdToRecord = new Dictionary<string, ClassRecord>();

            //班級名稱對照
            foreach (ClassRecord cr in classRecords)
            {
                if (!classIdToRecord.ContainsKey(cr.ID))
                {
                    classIdToRecord.Add(cr.ID, cr);
                }
            }

            //學生物件整理
            List<StudentObj> stuObjs = new List<StudentObj>();
            foreach (StudentRecord s in studentRecords)
            {
                StudentObj obj = new StudentObj(s);
                obj.ClassRecord = classIdToRecord.ContainsKey(s.RefClassID) ? classIdToRecord[s.RefClassID] : new ClassRecord();
                stuObjs.Add(obj);
            }

            stuObjs.Sort(delegate(StudentObj x, StudentObj y)
            {
                string x1 = x.ClassRecord.DisplayOrder.PadLeft(3, '0');
                string xx = (x.ClassRecord.GradeYear + "").PadLeft(3, '0');
                xx += x1 == "000" ? "999" : x1;
                xx += x.ClassRecord.Name.PadLeft(20, '0');
                xx += (x.StudentRecord.SeatNo + "").PadLeft(3, '0');

                string y1 = y.ClassRecord.DisplayOrder.PadLeft(3, '0');
                string yy = (y.ClassRecord.GradeYear + "").PadLeft(3, '0');
                yy += y1 == "000" ? "999" : y1;
                yy += y.ClassRecord.Name.PadLeft(20, '0');
                yy += (y.StudentRecord.SeatNo + "").PadLeft(3, '0');

                return xx.CompareTo(yy);
            });

            //領域及科目的聯集清單
            List<string> domains = new List<string>();
            List<string> subjects = new List<string>();

            Dictionary<string, Dictionary<string, DomainScore>> MakeUpDomainDic = new Dictionary<string, Dictionary<string, DomainScore>>();
            Dictionary<string, Dictionary<string, SubjectScore>> MakeUpSubjDic = new Dictionary<string, Dictionary<string, SubjectScore>>();

            foreach (JHSemesterScoreRecord JHssr in semesterScoreList)
            {
                //領域
                foreach (KeyValuePair<string, DomainScore> item in JHssr.Domains)
                {
                    //item.Value.Domain == item.Key;
                    //item.Value.Score
                    if (item.Value.Score.HasValue && item.Value.Score.Value < 60)
                    {
                        if (!MakeUpDomainDic.ContainsKey(JHssr.RefStudentID))
                            MakeUpDomainDic.Add(JHssr.RefStudentID, new Dictionary<string, DomainScore>());

                        MakeUpDomainDic[JHssr.RefStudentID].Add(item.Key, item.Value);
                        domains.Add(item.Key);
                    }
                }

                //科目
                foreach (string subj in JHssr.Subjects.Keys)
                {
                    SubjectScore ss = JHssr.Subjects[subj];
                    if (JHssr.Subjects[subj].Score.HasValue && JHssr.Subjects[subj].Score.Value < 60)
                    {
                        if (!MakeUpSubjDic.ContainsKey(JHssr.RefStudentID))
                            MakeUpSubjDic.Add(JHssr.RefStudentID, new Dictionary<string, SubjectScore>());

                        //科目名稱不應該有重覆
                        MakeUpSubjDic[JHssr.RefStudentID].Add(subj, ss);

                        //科目聯集清單
                        if (!subjects.Contains(subj))
                            subjects.Add(subj);
                    }
                }
            }

            //領域排序
            domains = domains.Distinct().ToList();
            domains.Sort(domainSort);

            //科目排序
            subjects.Sort(subjSort);

            //各領域位置的整理
            Dictionary<string, int> domainColumn = new Dictionary<string, int>();
            int index = 7;
            foreach (string key in domains)
            {
                if (!domainColumn.ContainsKey(key))
                    domainColumn.Add(key, index);
                index++;
            }

            //各科目位置的整理
            Dictionary<string, int> subjColumn = new Dictionary<string, int>();
            index = 7;
            foreach (string key in subjects)
            {
                if (!subjColumn.ContainsKey(key))
                    subjColumn.Add(key, index);
                index++;
            }

            //開始列印
            Workbook wb = new Workbook();
            wb.Open(new MemoryStream(Properties.Resources.Template));
            Worksheet templateSheet = wb.Worksheets["Template"];
            int sheet1 = wb.Worksheets.AddCopy("Template");
            int sheet2 = wb.Worksheets.AddCopy("Template");
            int sheet3 = wb.Worksheets.AddCopy("Template");

            wb.Worksheets[sheet1].Name = "領域補考清單";
            wb.Worksheets[sheet2].Name = "科目補考清單(橫)";
            wb.Worksheets[sheet3].Name = "科目補考清單(直)";

            //wb.Worksheets[sheet1].Cells[0, 0].PutValue("年級");
            //wb.Worksheets[sheet1].Cells[0, 1].PutValue("班級");
            //wb.Worksheets[sheet1].Cells[0, 2].PutValue("座號");
            //wb.Worksheets[sheet1].Cells[0, 3].PutValue("學號");
            //wb.Worksheets[sheet1].Cells[0, 4].PutValue("姓名");
            //wb.Worksheets[sheet1].Cells[0, 5].PutValue("學年度");
            //wb.Worksheets[sheet1].Cells[0, 6].PutValue("學期");

            wb.Worksheets[sheet1].SetColumnHeaders(new string[] { "年級", "班級", "座號", "學號", "姓名", "學年度", "學期"});

            //wb.Worksheets[sheet2].Cells[0, 0].PutValue("年級");
            //wb.Worksheets[sheet2].Cells[0, 1].PutValue("班級");
            //wb.Worksheets[sheet2].Cells[0, 2].PutValue("座號");
            //wb.Worksheets[sheet2].Cells[0, 3].PutValue("學號");
            //wb.Worksheets[sheet2].Cells[0, 4].PutValue("姓名");
            //wb.Worksheets[sheet2].Cells[0, 5].PutValue("學年度");
            //wb.Worksheets[sheet2].Cells[0, 6].PutValue("學期");

            wb.Worksheets[sheet2].SetColumnHeaders(new string[] { "年級", "班級", "座號", "學號", "姓名", "學年度", "學期" });

            //wb.Worksheets[sheet3].Cells[0, 0].PutValue("年級");
            //wb.Worksheets[sheet3].Cells[0, 1].PutValue("班級");
            //wb.Worksheets[sheet3].Cells[0, 2].PutValue("座號");
            //wb.Worksheets[sheet3].Cells[0, 3].PutValue("學號");
            //wb.Worksheets[sheet3].Cells[0, 4].PutValue("姓名");
            //wb.Worksheets[sheet3].Cells[0, 5].PutValue("學年度");
            //wb.Worksheets[sheet3].Cells[0, 6].PutValue("學期");
            //wb.Worksheets[sheet3].Cells[0, 7].PutValue("科目");
            //wb.Worksheets[sheet3].Cells[0, 8].PutValue("分數");

            wb.Worksheets[sheet3].SetColumnHeaders(new string[] { "年級", "班級", "座號", "學號", "姓名", "學年度", "學期", "科目", "分數" });

            //sheet1的domain標題
            int subIndex = 7;
            foreach (string d in domains)
            {
                wb.Worksheets[sheet1].Cells[0, subIndex].PutValue(d);
                subIndex++;
            }

            //sheet2的subject標題
            subIndex = 7;
            foreach (string s in subjects)
            {
                wb.Worksheets[sheet2].Cells[0, subIndex].PutValue(s);
                subIndex++;
            }

            int sheet1Row = 1;
            int sheet2Row = 1;
            int sheet3Row = 1;

            foreach (StudentObj so in stuObjs)
            {
                //只列印一般生
                if (so.StudentRecord.Status != StudentRecord.StudentStatus.一般)
                    continue;

                //sheet1
                if (MakeUpDomainDic.ContainsKey(so.StudentRecord.ID))
                {
                    wb.Worksheets[sheet1].Cells[sheet1Row, 0].PutValue(so.ClassRecord.GradeYear);
                    wb.Worksheets[sheet1].Cells[sheet1Row, 1].PutValue(so.ClassRecord.Name);
                    wb.Worksheets[sheet1].Cells[sheet1Row, 2].PutValue(so.StudentRecord.SeatNo);
                    wb.Worksheets[sheet1].Cells[sheet1Row, 3].PutValue(so.StudentRecord.StudentNumber);
                    wb.Worksheets[sheet1].Cells[sheet1Row, 4].PutValue(so.StudentRecord.Name);
                    wb.Worksheets[sheet1].Cells[sheet1Row, 5].PutValue(_schoolYear);
                    wb.Worksheets[sheet1].Cells[sheet1Row, 6].PutValue(_semester);

                    foreach (string d in domains)
                    {
                        int columnIndex = domainColumn[d];
                        if (MakeUpDomainDic[so.StudentRecord.ID].ContainsKey(d))
                        {
                            wb.Worksheets[sheet1].Cells[sheet1Row, columnIndex].PutValue(MakeUpDomainDic[so.StudentRecord.ID][d].Score);
                        }
                    }
                    sheet1Row++;
                }

                //sheet2跟sheet3一起處理
                if (MakeUpSubjDic.ContainsKey(so.StudentRecord.ID))
                {
                    wb.Worksheets[sheet2].Cells[sheet2Row, 0].PutValue(so.ClassRecord.GradeYear);
                    wb.Worksheets[sheet2].Cells[sheet2Row, 1].PutValue(so.ClassRecord.Name);
                    wb.Worksheets[sheet2].Cells[sheet2Row, 2].PutValue(so.StudentRecord.SeatNo);
                    wb.Worksheets[sheet2].Cells[sheet2Row, 3].PutValue(so.StudentRecord.StudentNumber);
                    wb.Worksheets[sheet2].Cells[sheet2Row, 4].PutValue(so.StudentRecord.Name);
                    wb.Worksheets[sheet2].Cells[sheet2Row, 5].PutValue(_schoolYear);
                    wb.Worksheets[sheet2].Cells[sheet2Row, 6].PutValue(_semester);

                    foreach (string s in subjects)
                    {
                        int columnIndex = subjColumn[s];
                        if (MakeUpSubjDic[so.StudentRecord.ID].ContainsKey(s))
                        {
                            wb.Worksheets[sheet2].Cells[sheet2Row, columnIndex].PutValue(MakeUpSubjDic[so.StudentRecord.ID][s].Score);
                        }
                    }
                    sheet2Row++;

                    //sheet3
                    foreach (string subj in MakeUpSubjDic[so.StudentRecord.ID].Keys)
                    {
                        wb.Worksheets[sheet3].Cells[sheet3Row, 0].PutValue(so.ClassRecord.GradeYear);
                        wb.Worksheets[sheet3].Cells[sheet3Row, 1].PutValue(so.ClassRecord.Name);
                        wb.Worksheets[sheet3].Cells[sheet3Row, 2].PutValue(so.StudentRecord.SeatNo);
                        wb.Worksheets[sheet3].Cells[sheet3Row, 3].PutValue(so.StudentRecord.StudentNumber);
                        wb.Worksheets[sheet3].Cells[sheet3Row, 4].PutValue(so.StudentRecord.Name);
                        wb.Worksheets[sheet3].Cells[sheet3Row, 5].PutValue(_schoolYear);
                        wb.Worksheets[sheet3].Cells[sheet3Row, 6].PutValue(_semester);

                        wb.Worksheets[sheet3].Cells[sheet3Row, 7].PutValue(subj);
                        wb.Worksheets[sheet3].Cells[sheet3Row, 8].PutValue(MakeUpSubjDic[so.StudentRecord.ID][subj].Score);

                        sheet3Row++;
                    }
                }
            }

            //total += sw.Elapsed.TotalMilliseconds; //計時標籤 40s
            //Console.WriteLine(total);

            wb.Worksheets[sheet1].AutoFitColumns();
            wb.Worksheets[sheet2].AutoFitColumns();
            wb.Worksheets[sheet3].AutoFitColumns();

            wb.Worksheets.RemoveAt(0);

            e.Result = wb;
        }

		private int domainSort(string x, string y)
		{
            if (_domainNameSort.Contains(x) && _domainNameSort.Contains(y))
			{
                return _domainNameSort.IndexOf(x).CompareTo(_domainNameSort.IndexOf(y));
			}
            else if (!_domainNameSort.Contains(x) && _domainNameSort.Contains(y))
			{
				return 1;
			}
            else if (_domainNameSort.Contains(x) && !_domainNameSort.Contains(y))
			{
				return -1;
			}
			return x.CompareTo(y);
		}

        private int subjSort(string x, string y)
        {
            if (_subjNameSort.Contains(x) && _subjNameSort.Contains(y))
            {
                return _subjNameSort.IndexOf(x).CompareTo(_subjNameSort.IndexOf(y));
            }
            else if (!_subjNameSort.Contains(x) && _subjNameSort.Contains(y))
            {
                return 1;
            }
            else if (_subjNameSort.Contains(x) && !_subjNameSort.Contains(y))
            {
                return -1;
            }
            return x.CompareTo(y);
        }

		private void btnPrint_Click(object sender, EventArgs e)
		{
            int sy, sm;
			if (!int.TryParse(cbbSchoolYear.Text, out sy))
			{
				MsgBox.Show("學年度必須選擇為數字");
				return;
			}
			if (!int.TryParse(cbbSemester.Text, out sm))
			{
				MsgBox.Show("學期必須選擇為數字");
				return;
			}

            string strClassId = string.Join(",", _classIds);
            if (strClassId == "")
            {
                MsgBox.Show("無選取班級，請確認是否選取班級");
                return;
            }

            if (_BW.IsBusy)
            {
                MsgBox.Show("系統忙碌中請稍後再試...");
                return;
            }
            else
            {
                btnPrint.Enabled = false;
                _schoolYear = sy;
                _semester = sm;

                _BW.RunWorkerAsync();
            }
		}

		private void btnQuit_Click(object sender, EventArgs e)
		{
			this.Close();
		}
	}

	public class StudentObj
	{
		public StudentRecord StudentRecord;
		public ClassRecord ClassRecord;
		public StudentObj(StudentRecord s)
		{
			this.StudentRecord = s;
		}
	}

    public static class Extends
    {
        public static void SetColumnHeaders(this Worksheet ws, params string[] ss)
        {
            int i = 0;

            foreach (string s in ss)
            {
                ws.Cells[0, i].PutValue(s);
                i++;
            }
        }
    }
}
