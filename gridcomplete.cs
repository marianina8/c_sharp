using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GridComplete
{
    public partial class Form1 : Form
    {
        private static Excel.Workbook MyBook;
        private static Excel.Application MyApp;
        private static Excel.Worksheet MySheet;


        private DateTime RoundUp(DateTime dt, TimeSpan d)
        {
            var delta = (d.Ticks - (dt.Ticks % d.Ticks)) % d.Ticks;
            return new DateTime(dt.Ticks + delta, dt.Kind);
        }

        private DateTime RoundDown(DateTime dt, TimeSpan d)
        {
            var delta = dt.Ticks % d.Ticks;
            return new DateTime(dt.Ticks - delta, dt.Kind);
        }

        private DateTime RoundToNearest(DateTime dt, TimeSpan d)
        {
            var delta = dt.Ticks % d.Ticks;
            bool roundUp = delta > d.Ticks / 2;

            return roundUp ? RoundUp(dt, d) : RoundDown(dt, d);
        }

        private int RTRound(int num)
        {
            int x = 0;
            int y = 8;

            if (num > 0 && num < 8)
                return 0;
            if(num==158 || num ==22)
            {
                x=x;
            }
            int count = 1;
            while(y < 500)
            {
                if (num >= x && num < y)
                    return (x+y)/ 2;

                x = y;
                if (count % 2 == 0)
                    y += 16;
                else
                    y += 14;

                count++;
            }

            return (x / y) / 2;
     
        }

        private Excel.Application newApp = null;
        private Excel.Workbook newWorkbook = null;
        private Excel.Worksheet newWorksheet = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void bOpenFileDialog2_Click(object sender, EventArgs e)
        {
            openFileDialog2.Filter = "Excel|*.xls|Excel 2010|*.xlsx";
            openFileDialog2.FilterIndex = 1;
            openFileDialog2.Multiselect = false;

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                tbFile2.Text = openFileDialog2.FileName;
            }
        }

        private void bOutputFolder_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                tbOutputFolder.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void bRun_Click(object sender, EventArgs e)
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(tbFile2.Text.ToString());
            MySheet = (Excel.Worksheet)MyBook.Sheets[1];
            DateTime sunday = new DateTime();
            double colWidth = 23.5;//18.17;
            double colWidthTime = 11;//6.69;
            int gProgress = 0;

            bOutputFolder.Enabled = false;
            bOpenFileDialog2.Enabled = false;
            bRun.Enabled = false;

            linkLabel1.Hide();
            progressBar1.Show();
            
            // Grab info from reports csv
            List<Telecast> telecasts = new List<Telecast>();
            
            gProgress = 5;
            progressBar1.Value = gProgress;

             gProgress = 10;
            progressBar1.Value = gProgress;

            try
            {
                for (int i = 11; i <= MySheet.Rows.Count; i++)
                {
                    Telecast t = new Telecast();
                    if (MySheet.Cells[i, 1].Value == null || String.IsNullOrWhiteSpace(MySheet.Cells[i, 1].Value.ToString()))
                        break;
                    t.myTlcst = MySheet.Cells[i, 1].Value.ToString();
                    if (MySheet.Cells[i, 2].Value != null)
                        t.myTitle = MySheet.Cells[i, 2].Value.ToString();
                   
                    t.myEST = RoundDown(DateTime.FromOADate(MySheet.Cells[i, 4].Value), TimeSpan.FromMinutes(15));
                   
                    t.myDay = MySheet.Cells[i, 3].Value;

                    DateTime start1 = new DateTime(t.myDay.Year, t.myDay.Month, t.myDay.Day, 0, 0, 0);
                    start1 = start1.AddHours(6);
                    DateTime end1 = new DateTime(t.myDay.Year, t.myDay.Month, t.myDay.Day, 0, 0, 0);
                    end1 = end1.AddHours(23);
                    end1 = end1.AddMinutes(59);

                    DateTime start2 = new DateTime(t.myDay.Year, t.myDay.Month, t.myDay.Day, 0, 0, 0);
                    DateTime end2 = new DateTime(t.myDay.Year, t.myDay.Month, t.myDay.Day, 0, 0, 0);
                    end2 = end2.AddHours(6);

                    if ((t.myEST >= start1) && (t.myEST <= end1))
                    {
                        //t.myDay = t.myDay.AddDays(1);
                    }
                    else if ((t.myEST >= start2) && (t.myEST <= end2))
                    {
                        t.myDay = t.myDay.AddDays(-1);
                    }
                    else
                    {
                        // do nothing
                    }
                    t.myTrt = MySheet.Cells[i, 5].Value.ToString();

                    t.myAA = Math.Round(MySheet.Cells[i, 6].Value).ToString();
                    telecasts.Add(t);
                }
            } catch (Exception e1)
            {
                return;
            }
            if(telecasts.Count ==0)
            {
                MessageBox.Show("ERROR: Check format of report spreadsheet.");
                return;
            }
            gProgress = 15;
            progressBar1.Value = gProgress;
           

            try {
            // generate grid spreadsheet
            newApp = new Excel.Application();
            newApp.Visible = false;
            newWorkbook = newApp.Workbooks.Add(1);
            newWorksheet = (Excel.Worksheet)newWorkbook.Sheets[1];
            DateTime startTime = new DateTime(2013, 9, 15, 0, 0, 0);
            newWorksheet.PageSetup.LeftHeader = "KEY:\nGreen cells ≥ +15% previous quarter's daypart average\nWhite cells (no color) b/t -15% and +15% versus previous quarter's daypart average\nYellow cells ≤ -15% previous quarter's daypart average";

            newWorksheet.PageSetup.LeftMargin = .20*72;
            newWorksheet.PageSetup.RightMargin = .20 * 72;
            newWorksheet.PageSetup.TopMargin = .82 * 72;
            newWorksheet.PageSetup.BottomMargin = .20 * 72;

            startTime = startTime.AddHours(6);
            newWorksheet.Cells[1, 1].EntireColumn.ColumnWidth = colWidthTime;
            Excel.Borders border = newWorksheet.Cells[1, 1].Borders;

            newWorksheet.Cells[2, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            border = newWorksheet.Cells[2, 1].Borders;
            border.LineStyle = Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            for (int i = 3; i <= 98; i++)
            {
                string lefttime = startTime.ToString("hh:mm tt");
                lefttime = lefttime.Remove(lefttime.Length - 1);
                newWorksheet.Cells[i, 1].NumberFormat = "@";
                newWorksheet.Cells[i, 1] = lefttime;
                newWorksheet.Cells[i, 1].EntireRow.RowHeight = 12;
                newWorksheet.Cells[i, 1].Font.Size = 9;
                newWorksheet.Cells[i, 1].Font.Bold = true;
                newWorksheet.Cells[i, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                newWorksheet.Cells[i, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                newWorksheet.Cells[i, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                border = newWorksheet.Cells[i, 1].Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;
                startTime = startTime.AddMinutes(15);

            }

            newWorksheet.Cells[99, 1].Value = "Total Day";
            newWorksheet.Cells[99, 1].EntireRow.RowHeight = 12;
            newWorksheet.Cells[99, 1].Font.Size = 9;
            newWorksheet.Cells[99, 1].Font.Bold = true;
            newWorksheet.Cells[99, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            newWorksheet.Cells[99, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            border = newWorksheet.Cells[99, 1].Borders;
            border.LineStyle = Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;
            
            gProgress = 30;
            progressBar1.Value = gProgress;
           

            var items = telecasts.Select(x => x.myDay).Distinct();
            sunday = items.Last();
            int j = 2;

            if(items.Count()==0)
            {
                MessageBox.Show("ERROR: No dates found.  Please check report file format.");
            }

            foreach (DateTime day in items.ToList())
            {

                    // set day header
                    newWorksheet.Cells[2, j].Value = day.ToString("M/dd ddd", CultureInfo.CreateSpecificCulture("en-US"));
                    newWorksheet.Cells[2, j].EntireColumn.ColumnWidth = colWidth;
                    newWorksheet.Cells[2, j].Font.Size = 12;
                    newWorksheet.Cells[2, j].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    newWorksheet.Cells[2, j].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    newWorksheet.Cells[2, j].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    newWorksheet.Cells[2, j].Font.Bold = true;

                    border = newWorksheet.Cells[2, j].Borders;
                    border.LineStyle = Excel.XlLineStyle.xlContinuous;
                    border.Weight = 2d;

                    // populate telecasts for that day
                    List<Telecast> subTelecasts = new List<Telecast>();
                    subTelecasts = telecasts.FindAll(
                        delegate(Telecast t)
                        {
                            return t.myDay == day;
                        }
                        );

                    for (int i = 3; i <= 98; i++)
                    {
                        Telecast subTelecast = new Telecast();
                        subTelecast = subTelecasts.Find(
                            delegate(Telecast t)
                            {
                                return t.myESTstring == newWorksheet.Cells[i, 1].Text;
                            }
                        );

                        if (subTelecast != null)
                        {
                            if(subTelecast.myTitle.Contains("Chr")){
                                int x =0;
                            }
                            int spaces_down = (RTRound(Convert.ToInt32(subTelecast.myTrt)) / 15);
                            string cellText = subTelecast.myTlcst.ToUpper().Trim() + " " + subTelecast.myTitle.Trim() + " " + String.Format("{0:#,##0}", Convert.ToInt32(subTelecast.myAA)) + " HH";
                            
                            if (cellText.Length > (colWidth * (spaces_down)))
                            {
                                newWorksheet.Cells[i, j].Font.Size = 9;
                                if (spaces_down == 4)
                                {
                                    if (newWorksheet.Cells[i, j].EntireRow.Height < 13)
                                    {
                                        newWorksheet.Cells[i, j].EntireRow.RowHeight = 13;
                                        newWorksheet.Cells[i + 1, j].EntireRow.RowHeight = 13;
                                        newWorksheet.Cells[i + 2, j].EntireRow.RowHeight = 13;
                                        newWorksheet.Cells[i + 3, j].EntireRow.RowHeight = 13;
                                    }
                                
                                }
                                if (spaces_down == 3)
                                {
                                    if (newWorksheet.Cells[i, j].EntireRow.Height < 14)
                                    {
                                        newWorksheet.Cells[i, j].EntireRow.RowHeight = 14;
                                        newWorksheet.Cells[i + 1, j].EntireRow.RowHeight = 14;
                                        newWorksheet.Cells[i + 2, j].EntireRow.RowHeight = 14;
                                    }
                                }
                                if(spaces_down == 2)
                                {
                                    if (newWorksheet.Cells[i, j].EntireRow.Height < 15)
                                    {
                                        newWorksheet.Cells[i, j].EntireRow.RowHeight = 15;
                                        newWorksheet.Cells[i + 1, j].EntireRow.RowHeight = 15;
                                    }
                                }
                                if(spaces_down == 1)
                                {
                                    if (newWorksheet.Cells[i, j].EntireRow.Height < 40)
                                    {
                                        newWorksheet.Cells[i, j].EntireRow.RowHeight = 36;
                                    }
                                }
                            }
                            else
                            {
                                newWorksheet.Cells[i, j].Font.Size = 9;
                            }

                            string sTextVal = (string)(newWorksheet.Cells[i-1,j] as Excel.Range).Value;
                            bool bMerged = newWorksheet.Cells[i - 1, j].MergeCells;

                            if(string.IsNullOrEmpty(sTextVal) && bMerged == false)
                            {
                                newWorksheet.Range[newWorksheet.Cells[i-1, j], newWorksheet.Cells[i, j]].Merge();
                                newWorksheet.Range[newWorksheet.Cells[i-1, j], newWorksheet.Cells[i + spaces_down - 1, j]].Merge();
                                newWorksheet.Range[newWorksheet.Cells[i-1, j], newWorksheet.Cells[i + spaces_down - 1, j]].Value = cellText;
                                newWorksheet.Range[newWorksheet.Cells[i-1, j], newWorksheet.Cells[i + spaces_down - 1, j]].WrapText = true;
                                newWorksheet.Range[newWorksheet.Cells[i-1, j], newWorksheet.Cells[i + spaces_down - 1, j]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                                border = newWorksheet.Range[newWorksheet.Cells[i-1, j], newWorksheet.Cells[i + spaces_down - 1, j]].Borders;
                                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                                border.Weight = 2d;
                            }
                            else if (newWorksheet.Cells[i, j].MergeCells)
                            {
                                i += 1;
                                newWorksheet.Range[newWorksheet.Cells[i, j], newWorksheet.Cells[i + spaces_down - 1, j]].Merge();
                                newWorksheet.Range[newWorksheet.Cells[i, j], newWorksheet.Cells[i + spaces_down - 1, j]].Value = cellText;
                                newWorksheet.Range[newWorksheet.Cells[i, j], newWorksheet.Cells[i + spaces_down - 1, j]].WrapText = true;
                                newWorksheet.Range[newWorksheet.Cells[i, j], newWorksheet.Cells[i + spaces_down - 1, j]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                                border = newWorksheet.Range[newWorksheet.Cells[i, j], newWorksheet.Cells[i + spaces_down - 1, j]].Borders;
                                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                                border.Weight = 2d;
                            }
                            else
                            {
                                newWorksheet.Range[newWorksheet.Cells[i, j], newWorksheet.Cells[i + spaces_down - 1, j]].Merge();
                                newWorksheet.Range[newWorksheet.Cells[i, j], newWorksheet.Cells[i + spaces_down - 1, j]].Value = cellText;
                                newWorksheet.Range[newWorksheet.Cells[i, j], newWorksheet.Cells[i + spaces_down - 1, j]].WrapText = true;
                                newWorksheet.Range[newWorksheet.Cells[i, j], newWorksheet.Cells[i + spaces_down - 1, j]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                                border = newWorksheet.Range[newWorksheet.Cells[i, j], newWorksheet.Cells[i + spaces_down - 1, j]].Borders;
                                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                                border.Weight = 2d;
                            }
                        }
                    }
                    
                    border = newWorksheet.Cells[99, j].Borders;
                    border.LineStyle = Excel.XlLineStyle.xlContinuous;
                    border.Weight = 2d;

                    j++;
                    gProgress += 10;
                    progressBar1.Value = gProgress;
                   
                }


                newWorksheet.Cells[2, j].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                border = newWorksheet.Cells[2, j].Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                for (int i = 3; i <= 98; i++ )
                {
                    border = newWorksheet.Cells[i, j].Borders;
                    border.LineStyle = Excel.XlLineStyle.xlContinuous;
                    border.Weight = 2d;

                    newWorksheet.Cells[i, j].Value = newWorksheet.Cells[i, 1].Text;
                    newWorksheet.Cells[i, j].Font.Size = 9;
                    newWorksheet.Cells[i, j].Font.Bold = true;
                    newWorksheet.Cells[i, j].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    newWorksheet.Cells[i, j].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    newWorksheet.Cells[i, j].EntireColumn.ColumnWidth = 9;
                    newWorksheet.Cells[i, j].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                }

                newWorksheet.Cells[99, j].EntireColumn.ColumnWidth = 9;
                newWorksheet.Cells[99, j].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                border = newWorksheet.Cells[99, j].Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                gProgress = 95;
                progressBar1.Value = gProgress;
               

            } 
            catch (Exception e2) {
                return;
            }


            try {
                if (File.Exists(tbOutputFolder.Text + "\\Weekly Viewership Flash Grid " + sunday.ToString("yyyy-MM-dd", CultureInfo.CreateSpecificCulture("en-US")) + ".xlsx"))
                {
                    File.Delete(tbOutputFolder.Text + "\\Weekly Viewership Flash Grid " + sunday.ToString("yyyy-MM-dd", CultureInfo.CreateSpecificCulture("en-US")) + ".xlsx");
                }
                newWorkbook.SaveAs(tbOutputFolder.Text + "\\Weekly Viewership Flash Grid " + sunday.ToString("yyyy-MM-dd", CultureInfo.CreateSpecificCulture("en-US")) + ".xlsx");
                newApp.Application.Quit();


                gProgress = 100;
                progressBar1.Value = gProgress;

                linkLabel1.Show();
                linkLabel1.AutoEllipsis = true;
                linkLabel1.Text = tbOutputFolder.Text + "\\Weekly Viewership Flash Grid " + sunday.ToString("yyyy-MM-dd", CultureInfo.CreateSpecificCulture("en-US")) + ".xlsx";
            

               
            } catch(Exception e3) {
                MessageBox.Show("ERROR: Error saving file.  Check if file is already open.");
            }

            if (!File.Exists(tbOutputFolder.Text + "\\Weekly Viewership Flash Grid " + sunday.ToString("yyyy-MM-dd", CultureInfo.CreateSpecificCulture("en-US")) + ".xlsx"))
            {
                MessageBox.Show("ERROR: Error saving file.  Check if file is already open.");
            }

            progressBar1.Hide();
            bOutputFolder.Enabled = true;
            bOpenFileDialog2.Enabled = true;
            bRun.Enabled = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            progressBar1.Hide();
            linkLabel1.Hide();
        }

        private void linkLabel1_Click(object sender, EventArgs e)
        {
            MyApp = new Excel.Application();
            MyApp.Visible = true;
            MyBook = MyApp.Workbooks.Open(linkLabel1.Text);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1];
            MyApp.ActiveWindow.View = Excel.XlWindowView.xlPageLayoutView;
            MyApp.ActiveWindow.Zoom = 150;
        }

        private void aboutGridCompleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("GridComplete\nVersion 1.0b\nDeveloper: Marian Montagnino\nmmontagnino@gmail.com");
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

    }

    public class Telecast
    {
        private string Tlcst = "";
        public string myTlcst { get { return Tlcst; } set { Tlcst = value; } }
        private string Title = "";
        public string myTitle { get { return Title; } set { Title = value; } }
        private DateTime Day = new DateTime();
        public DateTime myDay { get { return Day; } set { Day = value; } }
        private DateTime EST = new DateTime();
        public DateTime myEST { get { return EST; } set { EST = value; } }
        private string Trt = "";
        public string myTrt { get { return Trt; } set { Trt = value; } }
        private string AA = "";
        public string myAA { get { return AA; } set { AA = value; } }
        private string ESTstring = "";
        public string myESTstring
        {
            get
            {
                return EST.ToString("hh:mm tt", CultureInfo.CreateSpecificCulture("en-US")).Substring(0, 7);
            }
            set { ESTstring = value; }
        }
        private string Daystring = "";
        public string myDaystring
        {
            get
            {
                return Day.ToString("M/dd ddd", CultureInfo.CreateSpecificCulture("en-US"));
            }
            set { Daystring = value; }
        }

    }

}
