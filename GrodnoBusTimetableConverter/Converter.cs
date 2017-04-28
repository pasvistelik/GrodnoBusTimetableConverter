using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Collections.ObjectModel;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using TransportClasses;

namespace GrodnoBusTimetableConverter
{
    class Converter
    {
        //private List<Timetable>[] timetables = null;

        private double GetCellColor(Excel.Range cell)
        {
            try
            {
                
                for (int q = 1; q < cell.Characters.Count; q++) if (Regex.IsMatch(cell.Characters[q, 1].Text, "[0-9]")) return cell.Characters[q, 1].Font.Color;
            }
            catch
            {
                return cell.Font.Color;
                //MessageBox.Show(cell.Text);
            }
            throw new Exception("Ячейка не содержит время.");
        }

        private Converter(string filepath)
        {
            DateTime time = new DateTime(), prev_time = new DateTime(), start_time = new DateTime();
            time = DateTime.Now;
            start_time = DateTime.Now;
            prev_time = DateTime.Now;

            string transportNumber = null, transportName = null;
            List<Timetable>[] fullTable = new List<Timetable>[2];
            List<Timetable>[] fullTable_depo = new List<Timetable>[2];


            List<DayOfWeek> workingDays = new List<DayOfWeek>(new DayOfWeek[] { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday });
            List<DayOfWeek> weekDays = new List<DayOfWeek>(new DayOfWeek[] { DayOfWeek.Saturday, DayOfWeek.Sunday });
            List<DayOfWeek> sunDays = new List<DayOfWeek>(new DayOfWeek[] { DayOfWeek.Sunday });
            List<DayOfWeek> satDays = new List<DayOfWeek>(new DayOfWeek[] { DayOfWeek.Saturday });
            List<DayOfWeek> allDays = new List<DayOfWeek>(new DayOfWeek[] { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday });
            List<DayOfWeek> withoutMondayDays = new List<DayOfWeek>(new DayOfWeek[] { DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday });
            List<DayOfWeek> mondayDays = new List<DayOfWeek>(new DayOfWeek[] { DayOfWeek.Monday });
            List<DayOfWeek> withoutMondayAndWeekDays = new List<DayOfWeek>(new DayOfWeek[] { DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday });

            //Regex timePattern = new Regex(@"((\s)*([0-9]{1,2})(\*?)(\s)*)+");
            Regex fullTimePattern = new Regex(@"([0-9]{1,2})(\*)");
            Regex timePattern = new Regex(@"([0-9]{1,2})(\*?)");
            //timetables = new List<Timetable>[2];
            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(filepath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл

            //prev_time = time; time = DateTime.Now; Console.WriteLine("Старт: " + (time - prev_time).ToString() + ".  Total time: " + (time - start_time).ToString());


            for (int k = 2, n = ObjWorkBook.Sheets.Count; k <= n; k++)
            {
                transportNumber = null;
                transportName = null;

                fullTable[0] = new List<Timetable>();
                fullTable_depo[0] = new List<Timetable>();
                fullTable[1] = new List<Timetable>();
                fullTable_depo[1] = new List<Timetable>();


                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[k]; //получить k-ый лист
                //MessageBox.Show(ObjWorkSheet.Name);

                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
                int allRows = lastCell.Row, allCollumns = lastCell.Column;
                string[,] list = new string[allCollumns, allRows]; // массив значений с листа равен по размеру листу

                Excel.Range cell;

                string routeNum = ObjWorkSheet.Name;

                //StreamWriter SW = new StreamWriter(new FileStream(@"..\..\..\..\GrodnoBusTimetableConverterResults\file." + routeNum + ".json", FileMode.Create, FileAccess.Write));

                bool ok = true;

                int part = 0;
                bool needCheckPart = true;
                string oldRouteName = null;//, oldRouteNum = null;
                for (int r = 1; ok && r <= allCollumns; r++)
                {
                    for (int j = 0, cellsToNext; j < allRows; j++) // по всем строкам
                    {

                        cell = ObjWorkSheet.Cells[j + 1, r];

                        if (cell.Text == "А" || cell.Text == "A") // Нашли расписание данного транспорта для конкретной остановки.
                        {
                            //prev_time = time; time = DateTime.Now; Console.WriteLine("A start: " + (time - prev_time).ToString() + ".  Total time: " + (time - start_time).ToString());
                            ok = false;
                            // Определим, сколько строк занимает расписание данной остановки.
                            for (cellsToNext = 1; j + cellsToNext < allRows; cellsToNext++) if (ObjWorkSheet.Cells[j + 1 + cellsToNext, r].Text == "A") break;

                            string routeName = ObjWorkSheet.Cells[j + 1 + 2, r + 1].Text;

                            if (needCheckPart && oldRouteName != null && oldRouteName != routeName)
                            {
                                needCheckPart = false;
                                part = 1;
                                //oldRouteNum = routeNum;
                            }
                            oldRouteName = routeName;

                            string routeCurrentStationName = ObjWorkSheet.Cells[j + 1, r + 4].Text;

                            if (transportNumber == null) transportNumber = routeNum;
                            if (transportName == null) transportName = routeName;


                            string filename = "full." + routeNum + ".(" + routeName.Replace("\"", "") + ")." + routeCurrentStationName.Replace("\"", "") + ".json";
                            string filename_depo = "file." + routeNum + ".(" + routeName.Replace("\"", "") + ")." + routeCurrentStationName.Replace("\"", "") + ".depo.json";


                            DirectoryInfo myPath = Directory.CreateDirectory(@"..\..\..\..\GrodnoBusTimetableConverterResults\" + routeNum);
                            DirectoryInfo myPath_depo = Directory.CreateDirectory(@"..\..\..\..\GrodnoBusTimetableConverterResults\" + routeNum + "_depo");


                            StreamWriter SW = new StreamWriter(new FileStream(myPath.FullName + @"\" + filename, FileMode.Create, FileAccess.Write));
                            StreamWriter SW_depo = new StreamWriter(new FileStream(myPath_depo.FullName + @"\" + filename_depo, FileMode.Create, FileAccess.Write));



                            int startString = j + 1 + 4;
                            int endString = j + cellsToNext;
                            int startColumn = r + 1;
                            int endColumn = r + 24;

                            //MessageBox.Show("Num: " + routeNum + "\nName: " + routeName + "\nStation: " + routeCurrentStationName + "\nСтроки: " + startString + " - " + endString + "\nСтолбцы: " + startColumn + " - " + endColumn);

                            Dictionary<double, int> numsTmp = new Dictionary<double, int>(3);
                            Queue<int> my_nums = new Queue<int>();
                            Excel.Range currentCell = null;
                            List<string> types = new List<string>();
                            StringBuilder type = new StringBuilder();
                            for (int strIndex = startString; strIndex <= endString; strIndex++)
                            {
                                bool needCheck = true;
                                for (int clmnIndex = startColumn; clmnIndex <= endColumn; clmnIndex++)
                                {
                                    currentCell = ObjWorkSheet.Cells[strIndex, clmnIndex];
                                    if (timePattern.IsMatch(currentCell.Text))
                                    {
                                        if (needCheck)
                                        {
                                            double color = GetCellColor(currentCell);
                                            if (!numsTmp.ContainsKey(color))
                                            {
                                                numsTmp.Add(color, strIndex);
                                                if (numsTmp.Count > 1)
                                                {
                                                    my_nums.Enqueue(strIndex);
                                                    //MessageBox.Show(type.ToString());
                                                    types.Add(type.ToString());
                                                    type = new StringBuilder();
                                                }
                                            }
                                            //MessageBox.Show("Ячейка [" + strIndex + ", " + clmnIndex + "] = " + currentCell.Text + " имеет цвет " + color);
                                            needCheck = false;
                                            //break;
                                        }
                                    }
                                    else
                                    {
                                        type.Append(currentCell.Text + " ");
                                    }
                                }
                            }
                            //MessageBox.Show(type.ToString());
                            types.Add(type.ToString());


                            //if (!numsTmp.ContainsValue(endString)) numsTmp.Add(-1, endString);
                            my_nums.Enqueue(endString + 1);

                            int[] colorsValues = numsTmp.Values.ToArray();
                            ObservableCollection<Table> t = new ObservableCollection<Table>();
                            ObservableCollection<Table> t2 = new ObservableCollection<Table>();

                            HashSet<DayOfWeek> usedDays = new HashSet<DayOfWeek>();

                            for (int start = startString, i = 0, end; i < numsTmp.Count; start = end, i++)
                            {
                                end = my_nums.Dequeue();
                                ObservableCollection<SimpleTime> tmpTimes = new ObservableCollection<SimpleTime>();
                                ObservableCollection<SimpleTime> tmpTimesToDepo = new ObservableCollection<SimpleTime>();
                                for (int clmnIndex = startColumn, hour = 5; clmnIndex <= endColumn; clmnIndex++, hour = (hour + 1) % 24)
                                {
                                    for (int strIndex = start; strIndex < end; strIndex++)
                                    {
                                        currentCell = ObjWorkSheet.Cells[strIndex, clmnIndex];
                                        //...
                                        foreach (Match s in timePattern.Matches(currentCell.Text))
                                        {
                                            string val = s.Value;
                                            //MessageBox.Show(val);
                                            if (val.Contains("*"))
                                            {
                                                tmpTimesToDepo.Add(new SimpleTime(hour, int.Parse(val.Replace("*", ""))));
                                                //MessageBox.Show(tmpTimesToDepo[tmpTimesToDepo.Count - 1].ToString());
                                            }
                                            else
                                            {
                                                tmpTimes.Add(new SimpleTime(hour, int.Parse(val)));
                                                //MessageBox.Show(tmpTimes[tmpTimes.Count-1].ToString());
                                            }
                                        }
                                    }
                                }
                                List<DayOfWeek> inner_days = null;
                                string my_days = types[i];
                                int my_color = colorsValues[i];
                                if (my_days.Contains("раб."))
                                {
                                    inner_days = workingDays;
                                    foreach (DayOfWeek day in inner_days) usedDays.Add(day);
                                }
                                else if (my_days.Contains("вых."))
                                {
                                    inner_days = weekDays;
                                    foreach (DayOfWeek day in inner_days) usedDays.Add(day);
                                }
                                else if (my_days.Contains("ежедневно"))
                                {
                                    inner_days = allDays;
                                    foreach (DayOfWeek day in inner_days) usedDays.Add(day);
                                }
                                else if (my_days.Contains("субб"))
                                {
                                    inner_days = satDays;
                                    foreach (DayOfWeek day in inner_days) usedDays.Add(day);
                                }
                                else if (my_days.Contains("воскр"))
                                {
                                    inner_days = sunDays;
                                    foreach (DayOfWeek day in inner_days) usedDays.Add(day);
                                }
                                else if (my_days.Contains("вт,ср,чт,пт,вых"))
                                {
                                    inner_days = withoutMondayDays;
                                    foreach (DayOfWeek day in inner_days) usedDays.Add(day);
                                }
                                else if (my_days.Contains("пн"))
                                {
                                    inner_days = mondayDays;
                                    foreach (DayOfWeek day in inner_days) usedDays.Add(day);
                                }
                                else if (my_days.Contains("вт,ср,чт,пт"))
                                {
                                    inner_days = withoutMondayAndWeekDays;
                                    foreach (DayOfWeek day in inner_days) usedDays.Add(day);
                                }
                                else if (my_color == 0)
                                {
                                    inner_days = workingDays;
                                    foreach (DayOfWeek day in inner_days) usedDays.Add(day);
                                }
                                else if (my_color == 32768)
                                {
                                    inner_days = satDays;
                                    foreach (DayOfWeek day in inner_days) usedDays.Add(day);
                                }
                                else if (my_color == 255)// || my_color == 98 || my_color == 105 || my_color == 9 || my_color == 5
                                {
                                    inner_days = weekDays;
                                    foreach (DayOfWeek day in inner_days) usedDays.Add(day);
                                }
                                else if (!(usedDays.Contains(DayOfWeek.Monday) && usedDays.Contains(DayOfWeek.Tuesday) && usedDays.Contains(DayOfWeek.Wednesday) && usedDays.Contains(DayOfWeek.Thursday) && usedDays.Contains(DayOfWeek.Friday)))
                                {
                                    inner_days = workingDays;
                                    foreach (DayOfWeek day in inner_days) usedDays.Add(day);
                                }
                                else if (!(usedDays.Contains(DayOfWeek.Saturday) && usedDays.Contains(DayOfWeek.Sunday)))
                                {
                                    inner_days = weekDays;
                                    foreach (DayOfWeek day in inner_days) usedDays.Add(day);
                                }
                                else
                                    MessageBox.Show(my_color.ToString());//inner_days = workingDays;

                                Table tmpTable = new Table(inner_days, tmpTimes);
                                Table tmpTableToDepo = new Table(inner_days, tmpTimesToDepo);

                                
                                t.Add(tmpTable);
                                t2.Add(tmpTableToDepo);
                            }

                            Timetable tbl = new Timetable(TableType.table, t);
                            fullTable[part].Add(tbl);
                            

                            Timetable tbl2 = new Timetable(TableType.table, t2);
                            fullTable_depo[part].Add(tbl2);

                            Console.WriteLine("Bus № " + routeNum + ":  ["+routeCurrentStationName+ "]   on   [" + routeName + "]  converted.");

                            SW.Write(tbl.Serialize());
                            SW.Close();
                            SW_depo.Write(tbl2.Serialize());
                            SW_depo.Close();

                            //prev_time = time; time = DateTime.Now; Console.WriteLine("A end: " + (time - prev_time).ToString() + ".  Total time: " + (time - start_time).ToString());
                        }
                    }
                }

                if (fullTable[0].Count != 0 && fullTable[1].Count != 0)
                {
                    Timetable firstTimetableStation = fullTable[0][0];
                    Timetable secondTimetableStation = fullTable[1][0];//fullTable[1].Count - 1
                    fullTable[0].Add(secondTimetableStation);
                    fullTable[1].Add(firstTimetableStation);
                }
                
                DirectoryInfo fullTablePath = Directory.CreateDirectory(@"..\..\..\..\GrodnoBusTimetableConverterResults");
                string fullTableFilename = "full.bus." + transportNumber + ".(" + transportName.Replace("\"", "") + ").json";
                string fullTableFilename_depo = "full.bus." + transportNumber + ".(" + transportName.Replace("\"", "") + ").depo.json";

                StreamWriter fullTableSW = new StreamWriter(new FileStream(fullTablePath.FullName + @"\" + fullTableFilename, FileMode.Create, FileAccess.Write));
                StreamWriter fullTableSW_depo = new StreamWriter(new FileStream(fullTablePath.FullName + @"\" + fullTableFilename_depo, FileMode.Create, FileAccess.Write));

                fullTableSW.Write(Timetable.SerializeFullTable(fullTable));
                fullTableSW.Close();

                fullTableSW_depo.Write(Timetable.SerializeFullTable(fullTable_depo));
                fullTableSW_depo.Close();

            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            Marshal.ReleaseComObject(ObjWorkBook);
            Marshal.ReleaseComObject(ObjWorkExcel);
            ObjWorkBook = null;
            ObjWorkExcel = null;
            GC.Collect(); // убрать за собой

        }
        
        public static void Convert(string filepath)
        {
            Converter converter = new Converter(filepath);
            //converter.FindDepoRoutes();

            //   return null;
        }
    }
}
