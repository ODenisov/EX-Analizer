using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace EX_Analizer
{
    public partial class Form1 : Form
    {
              public Form1()
        {
            InitializeComponent();
            FillMe();
        }

        private static void FillMe()
        {
            FirstCol.Add("Ежедневное ТО выполнено", 3);
            FirstCol.Add("Количество депозитов", 4);
            FirstCol.Add("Количество банкнот", 5);
            FirstCol.Add("Процент от теоретической выработки", 6);
            FirstCol.Add("Зафиксировано ошибок", 7);
            FirstCol.Add("Самые часто встречающиеся ошибки за смену ", 8);
            FirstCol.Add("Количество заходов внутрь комплекса(не чек-листы)", 11);
            FirstCol.Add("Время открытия смены", 12);
            FirstCol.Add("Время закрытия смены", 13);
            FirstCol.Add("Время открытия смены(не чек-листы)", 14);
            FirstCol.Add("Время закрытия смены(не чек-листы)", 15);
            FirstCol.Add("Время работы смены", 16);
            FirstCol.Add("Теоритический средний простой", 17);
            FirstCol.Add("Теоретический средний простой(не чек-листы)", 18);
            FirstCol.Add("Время ведения пересчёта внутри комплекса(не чек - листы)", 19);
            FirstCol.Add("Простой без загрузки(не чек-листы)", 20);
            FirstCol.Add("Производительность", 21);
            FirstCol.Add("Скорость", 22);
            FirstCol.Add("Мощность", 23);
        }

        static Dictionary<string, int> DateCols = new Dictionary<string, int>();

        private void FillDictCountTable(string _datefromxlsx)
        {
            if (DateCols.ContainsKey(_datefromxlsx))
            {
                return;
            }
            DateTime date = new DateTime(2018, 2, 1);
            for (int i = 0; i < 31; i++)
            {
                DateCols.Add(date.ToShortDateString(), i + 2);
                date = date.AddDays(1);
            }
        }

        ExcelWorksheet OpenInWorksheet(ExcelWorkbook _input,string _name)
        {
            if (_input.Worksheets[_name]==null)
            {
                return _input.Worksheets.Add(_name);
            }
            return _input.Worksheets[_name];
        }

        private string GetThatDate(object _date)
        {
            if (_date.GetType() == DateTime.Now.GetType())
            {
                return ((DateTime)_date).ToShortDateString();
            }
            else
            {
                return DateTime.FromOADate((double)_date).ToShortDateString();
            }
        }

        private string GetThatCity(ExcelWorksheet _input)
        {
            if (_input.Cells["H1"].Value == null)
            {
                return _input.Cells["H4"].Value == null ? " " : _input.Cells["H4"].Value.ToString().Split(' ').Last();
            }
            return _input.Cells["H1"].Value == null ? " " : _input.Cells["H1"].Value.ToString().Split(' ').Last();
        }

        int GetMoneyCount(ExcelWorksheet _input)
        {
            if (_input.Cells["H1"].Value == null)
            {
                return _input.Cells["J4"].Value == null ? 0 : Convert.ToInt32(_input.Cells["J4"].Value.ToString());
            }
            return _input.Cells["I4"].Value == null ? 0 : Convert.ToInt32(_input.Cells["I4"].Value.ToString());
        }

        int GetDepositCount(ExcelWorksheet _input)
        {
            if (_input.Cells["H1"].Value == null)
            {
                return _input.Cells["K4"].Value == null ? 0 : Convert.ToInt32(_input.Cells["K4"].Value.ToString());
            }
            return _input.Cells["J4"].Value == null ? 0 : Convert.ToInt32(_input.Cells["J4"].Value.ToString());
        }

        string GetThatTO(ExcelWorksheet _input)
        {
            if (_input.Cells["H1"].Value == null)
            {
                return _input.Cells["I4"].Value == null ? "Net" :_input.Cells["I4"].Value.ToString();
            }
            return _input.Cells["H4"].Value == null ? "Net" : _input.Cells["H4"].Value.ToString();
        }

        DateTime GetTimeRight(string _time)
        {
            string[] time = _time.Split(new char[] { ':', '.', ',', ';','-' });
            int hours;
            int minutes;
            if (time.Length < 2)
            {
                hours = 0;
                minutes = 0;
            }
            else
            {
                hours = Convert.ToInt32(time[0]);
                minutes = Convert.ToInt32(time[1]);
            }

            return new DateTime(2018, 4, 4, hours, minutes, 0);
        }

        private void MakeOut(string _filename)
        {
            ExcelPackage inputPackage = new ExcelPackage(new FileInfo(_filename));
            ExcelPackage package = new ExcelPackage(new FileInfo("outfile.xlsx"));

            var i_workbook = inputPackage.Workbook;
            var i_worksheet = i_workbook.Worksheets[1];
            var o_workbook = package.Workbook;

            var Total = OpenInWorksheet(package.Workbook,"Total");

            Dictionary<string,int> errorcount = new Dictionary<string, int>();
            List<double> errortime = new List<double>();

            int step = 2;
            for (int i = 7; i <= 37; i++)
            {
                errorcount.Add(i_worksheet.Cells[i, 1].Value.ToString(),
                               i_worksheet.Cells[i, 6].Value == null ?
                               0 :
                               Convert.ToInt32(i_worksheet.Cells[i, 6].Value.ToString())
                               );

                if (i_worksheet.Cells[i, 7].Value == null)
                {
                    errortime.Add(0);
                }
                else
                {
                    if (i_worksheet.Cells[i, 7].Value.GetType() == DateTime.Now.GetType())
                    {
                        DateTime temp = (DateTime)i_worksheet.Cells[i, 7].Value;
                        double hellothere = temp.ToOADate() * errorcount.Values.ToArray()[i - 7];
                        errortime.Add(hellothere);

                    }
                    else if (i_worksheet.Cells[i, 7].Value.GetType() == Double.MinValue.GetType())
                    {
                        double generalkenobi = (double)i_worksheet.Cells[i, 7].Value * errorcount.Values.ToArray()[i - 7];
                        errortime.Add(generalkenobi);
                    }
                    else
                    {
                        errortime.Add(0);
                    }
                }
            }//for

            double TotalStunTime = errortime.Sum();

            string i_date = GetThatDate(i_worksheet.Cells["B4"].Value);
            FillDictCountTable(i_date);
            string i_RUP = i_worksheet.Cells["F4"].Value.ToString() + " " + GetThatCity(i_worksheet);

            ExcelWorksheet o_worksheet;
            if (o_workbook.Worksheets[i_RUP] == null)
            {
                o_worksheet = o_workbook.Worksheets.Add(i_RUP);
                step = 0;
                foreach (string date in DateCols.Keys)
                {
                    o_worksheet.Cells[1, DateCols[date]+step].Value = date;
                    o_worksheet.Cells[2, DateCols[date] + step].Value = "Дневная смена";
                    o_worksheet.Cells[2, DateCols[date] + step + 1].Value = "Ночная смена";
                    o_worksheet.Column(DateCols[date] + step).AutoFit();
                    o_worksheet.Column(DateCols[date] + step+1).AutoFit();
                    o_worksheet.Cells[1, DateCols[date] + step, 1, DateCols[date] + step + 1].Merge = true;
                    step++;
                }
                foreach(string key in FirstCol.Keys)
                {
                    o_worksheet.Cells[FirstCol[key], 1].Value = key;
                }
                o_worksheet.Column(1).AutoFit();
                o_worksheet.Cells[8, 1, 10, 1].Merge = true;


            }
            else
            {
                o_worksheet = o_workbook.Worksheets[i_RUP];
            }
            //Fillings
            int dateIndexAM = DateCols[i_date] * 2 - 2;
            int dateIndexPM = DateCols[i_date] * 2 - 1;

            if (i_worksheet.Cells["D4"].Value != null)
            {
                if (i_worksheet.Cells["C4"].Value.ToString() == "Дневная")
                {
                    o_worksheet.Cells[3, dateIndexAM].Value = GetThatTO(i_worksheet);//ТО
                    o_worksheet.Cells[4, dateIndexAM].Value = GetDepositCount(i_worksheet);//Депозиты
                    o_worksheet.Cells[5, dateIndexAM].Value = GetMoneyCount(i_worksheet);//Купюры
                    int StunSum = 0;
                    foreach (string key in errorcount.Keys)
                    {
                        StunSum += errorcount[key];
                    }
                    o_worksheet.Cells[7, dateIndexAM].Value = StunSum;//кол-во ошибок

                    errorcount = errorcount.OrderByDescending(x => x.Value).ToDictionary(x => x.Key, x => x.Value);
                    string[] top = errorcount.Keys.ToArray();
                    o_worksheet.Cells[8, dateIndexAM].Value = "1. " + top[0];
                    o_worksheet.Cells[9, dateIndexAM].Value = "2. " + top[1];
                    o_worksheet.Cells[10, dateIndexAM].Value = "3. " + top[2];//Топ ошибок

                    o_worksheet.Row(8).Height = 50;
                    o_worksheet.Cells[12, dateIndexAM].Value = GetTimeRight(i_worksheet.Cells["D4"].Value.ToString());//Открытие смены
                    o_worksheet.Cells[12, dateIndexAM].Style.Numberformat.Format = "hh:mm";
                    if(i_worksheet.Cells["E4"].Value==null)
                    {
                        i_worksheet.Cells["E4"].Value = " ";
                    }
                    o_worksheet.Cells[13, dateIndexAM].Value = GetTimeRight(i_worksheet.Cells["E4"].Value.ToString()).AddDays(1);//Закрытие
                    o_worksheet.Cells[13, dateIndexAM].Style.Numberformat.Format = "hh:mm";

                    o_worksheet.Cells[16, dateIndexAM].Formula = o_worksheet.Cells[13, dateIndexAM].ToString() +
                                                          "-" +
                                                          o_worksheet.Cells[12, dateIndexAM].ToString();
                    o_worksheet.Cells[16, dateIndexAM].Style.Numberformat.Format = "hh:mm";//время работы смены

                    o_worksheet.Cells[17, dateIndexAM].Value = DateTime.FromOADate(TotalStunTime);
                    o_worksheet.Cells[17, dateIndexAM].Style.Numberformat.Format = "hh:mm";//время простоя(теор.)


                }
                else if (i_worksheet.Cells["C4"].Value.ToString() == "Ночная")
                {
                    o_worksheet.Cells[3, dateIndexPM].Value = GetThatTO(i_worksheet);//ТО
                    o_worksheet.Cells[4, dateIndexPM].Value = GetDepositCount(i_worksheet);//Депозиты
                    o_worksheet.Cells[5, dateIndexPM].Value = GetMoneyCount(i_worksheet);//Купюры
                    int StunSum = 0;
                    foreach (string key in errorcount.Keys)
                    {
                        StunSum += errorcount[key];
                    }
                    o_worksheet.Cells[7, dateIndexPM].Value = StunSum;//кол-во ошибок

                    errorcount = errorcount.OrderByDescending(x => x.Value).ToDictionary(x => x.Key, x => x.Value);
                    string[] top = errorcount.Keys.ToArray();
                    o_worksheet.Cells[8, dateIndexPM].Value = "1. " + top[0];
                    o_worksheet.Cells[9, dateIndexPM].Value = "2. " + top[1];
                    o_worksheet.Cells[10, dateIndexPM].Value = "3. " + top[2];//Топ ошибок

                    o_worksheet.Row(8).Height = 50;
                    o_worksheet.Cells[12, dateIndexPM].Value = GetTimeRight(i_worksheet.Cells["D4"].Value.ToString());//Открытие смены
                    o_worksheet.Cells[12, dateIndexPM].Style.Numberformat.Format = "hh:mm";
                    if (i_worksheet.Cells["E4"].Value == null)
                    {
                        i_worksheet.Cells["E4"].Value = " ";
                    }
                    o_worksheet.Cells[13, dateIndexPM].Value = GetTimeRight(i_worksheet.Cells["E4"].Value.ToString()).AddDays(1);//Закрытие
                    o_worksheet.Cells[13, dateIndexPM].Style.Numberformat.Format = "hh:mm";

                    o_worksheet.Cells[16, dateIndexPM].Formula = o_worksheet.Cells[13, dateIndexPM].ToString() +
                                                          "-" +
                                                          o_worksheet.Cells[12, dateIndexPM].ToString();

                    o_worksheet.Cells[16, dateIndexPM].Style.Numberformat.Format = "hh:mm";
                    o_worksheet.Cells[17, dateIndexPM].Value = DateTime.FromOADate(TotalStunTime);
                    o_worksheet.Cells[17, dateIndexPM].Style.Numberformat.Format = "hh:mm";//время простоя(теор.)
                }
            }
            package.Save();
            package.Dispose();
            inputPackage.Dispose();
        }

        private void b_openInput_Click(object sender, EventArgs e)
        {
            openExcelFile.ShowDialog();
        }

        private void b_startProcess_Click(object sender, EventArgs e)
        {
            if (openExcelFile.FileNames.Length < 2)
            {
                MakeOut(openExcelFile.FileName);
            }
            else
            {
                foreach (string file in openExcelFile.FileNames)
                {
                    MakeOut(file);
                }
            }
            MessageBox.Show("DONE");
        }

        private void openExcelFile_FileOk(object sender, CancelEventArgs e)
        {
            label1.Text = openExcelFile.FileName;
        }
    }
}
