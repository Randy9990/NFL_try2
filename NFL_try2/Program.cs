using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace NFL_try2
{
    class Program
    {
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static Excel.Range xlRange = null;

        static void Main(string[] args)
        {
           DateTime start = DateTime.Now;

            MyApp = new Excel.Application
            {
                Visible = false
            };
            string XLS_PATH = "C:\\Users\\randy.mccombs\\source\\repos\\NFL_try2\\NFL_try2\\NFL_Small_Set.xlsx";

            MyBook = MyApp.Workbooks.Open(XLS_PATH);
            MySheet = (Excel.Worksheet)MyBook.Sheets["Sheet1"];
            xlRange = MySheet.UsedRange;
            int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            SortedList<string, PlayerSeason> runningbacksList = new SortedList<string, PlayerSeason>();
            SortedList<string, PlayerSeason> quarterbacksList = new SortedList<string, PlayerSeason>();

            for (int x = 1; x <= lastRow; x++)
            {
                string position = xlRange.Cells.Value2[x, 22].ToString();

               
                if ((position == "RB") || (position == "QB"))
                {
                    PlayerSeason currentrow = new PlayerSeason();
                    currentrow.Player = xlRange.Cells.Value2[x, 2].ToString();
                    currentrow.PassingYards = xlRange.Cells.Value2[x, 11];
                    currentrow.RushingYards = xlRange.Cells.Value2[x, 15];
                    currentrow.Position = xlRange.Cells.Value2[x, 22].ToString();

                    if (position == "RB")
                    {

                        if (runningbacksList.ContainsKey(currentrow.Player))
                        {
                            runningbacksList[currentrow.Player].RushingYards += currentrow.RushingYards;
                            runningbacksList[currentrow.Player].PassingYards += currentrow.PassingYards;
                        }
                        else
                        {
                            runningbacksList.Add(currentrow.Player, currentrow);
                        }

                    }

                    if (position == "QB")
                    {

                        if (quarterbacksList.ContainsKey(currentrow.Player))
                        {
                            quarterbacksList[currentrow.Player].RushingYards += currentrow.RushingYards;
                            quarterbacksList[currentrow.Player].PassingYards += currentrow.PassingYards;
                        }
                        else
                        {
                            quarterbacksList.Add(currentrow.Player, currentrow);
                        }

                        //Console.WriteLine("player {0}, position {1} ", currentrow.Player, currentrow.Position);
                    }
                }

            }

            // for debugging:
            //foreach (string rbs in runningbacksList.Keys)
            //{
            //    Console.WriteLine("player \t{0}, \t \t position {1}, yards {2:0,0} ", rbs, runningbacksList[rbs].Position, runningbacksList[rbs].RushingYards);

            //}

            string bestRB = runningbacksList.OrderByDescending(s => s.Value.RushingYards).First().Key;
            Console.WriteLine("The best running back is \t{0}, yards {1:0,0} ", bestRB, runningbacksList[bestRB].RushingYards);

            string bestQB = quarterbacksList.OrderByDescending(s => s.Value.PassingYards).First().Key;
            Console.WriteLine("The best quarter back is \t{0}, yards {1:0,0} ", bestQB, quarterbacksList[bestQB].PassingYards);

            TimeSpan timeItTook = DateTime.Now - start;
            Console.WriteLine("elapsed time: \t{0}", timeItTook);


            Console.ReadKey(); 

            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(MySheet);
            MyBook.Close();
            Marshal.ReleaseComObject(MyBook);
            MyApp.Quit();
            Marshal.ReleaseComObject(MyApp);
        }

        //private static void BuildListRB(SortedList<string, PlayerSeason> runningbacksList, PlayerSeason currentrow)
        //{
        //    throw new NotImplementedException();
        //}

    }

}
