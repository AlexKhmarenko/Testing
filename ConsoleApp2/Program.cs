using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            
            
            string adress = $"{Environment.CurrentDirectory}\\Questions\\Questions.xlsx";

            Excel.Application xlApp = new();

            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;

            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(adress);
            int count = 0;
            int countQuestions = 1;
            string temp;
            string[] array = new string[5]; //Максимальное количество ответов
            int levelsQuantity;     //Количество уровней сложности
            int questionsQuontyty = 0; //Количество вопросов
            string userName;        //Юзер
            int level;              //Уровень сложности

            levelsQuantity = xlWorkbook.Sheets.Count;

            Console.WriteLine("ТЕСТ\n\n");
            Console.Write("Введите свои Имя и Фамилию: ");
            userName = Console.ReadLine();
            Console.Write("Введите уровень сложности: ");
            while (true) //Проверка ввода уровня сложности
            {
                string text = Console.ReadLine();
                if (int.TryParse(text, out int number))
                {
                    int key = Convert.ToInt32(text);
                    if (key > 0 && key <= levelsQuantity)
                    {
                        level = Convert.ToInt32(text);
                        break;
                    }
                    else
                    {
                        Console.Write("Такой вариант отсутствует, повторите ввод от 1 до {0}: ", (levelsQuantity));
                    }
                }
                else
                {
                    Console.Write("Необходимо ввести цифру от 1 до {0}: ", (levelsQuantity));
                }
            }
            Console.Clear();
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[level];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count; //Строки

            rowCount = xlRange.Rows.Count;
            for (int i = 1; i <= rowCount + 1; i++)
            {
                if (xlRange.Cells[i, 2] == null || xlRange.Cells[i, 2].Value2 == null)
                {
                    questionsQuontyty++;
                }
            }
            string[,] report = new string[questionsQuontyty, 5];//0 - Вопрос
                                                                //1 - Верный ответ
                                                                //2 - Ответ Юзера
                                                                //3 - Бал
                                                                //4 - Время на ответ


            for (int i = 1; i <= rowCount + 1; i++)
            {
                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
                {
                    temp = xlRange.Cells[i, 2].Value2;
                    array[count] = temp;
                    count++;
                }
                else
                {
                    Questions(array, countQuestions, count, questionsQuontyty, report);
                    count = 0;
                    countQuestions++;

                    array = new string[5];
                    Console.Clear();
                }
            }

            //xlWorkbook.Save();

            CloseExcel(xlRange, xlWorkbook, xlWorksheet, xlApp);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            

            ReportToExcel(userName, level, report, questionsQuontyty);

            GC.Collect();                           
            GC.WaitForPendingFinalizers();

            Console.CursorVisible = false;
            Console.WriteLine("Данные сохранены.\nНажмите <Enter> для завершения программы");
            
            Console.Read();


        }
        static void CloseExcel(Excel.Range xlRange, Excel.Workbook WorkBook, Excel.Worksheet WorkSheet, Excel.Application Excel)
        {
            //GC.Collect();
            //GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);

            Marshal.ReleaseComObject(WorkSheet);
            WorkBook.Close();
            Marshal.ReleaseComObject(WorkBook);
            Excel.Quit();
            Marshal.ReleaseComObject(Excel);

        }


        static void ReportToExcel(string userName, int level, string[,] report, int qq)
        {
            int score = 0;
            Excel.Application excel = new ();

            excel.Visible = false;
            excel.DisplayAlerts = false;

            Excel.Workbook workBook = excel.Workbooks.Add(Type.Missing);
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.ActiveSheet;

            workSheet.Name = userName;

            Excel.Range xlRange = workSheet.UsedRange;

            workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, 8]].Merge();
            workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[2, 8]].Merge();
            workSheet.Range[workSheet.Cells[3, 1], workSheet.Cells[3, 8]].Merge();
            workSheet.Cells[1, 1] = $"Результаты тестирования пользователя {userName}";
            workSheet.Cells[2, 1] = $"Время тестирования: {DateTime.Now.ToString()}";
            workSheet.Cells[3, 1] = $"Уровень сложности: {level}";
            workSheet.Cells[5, 1] = "№";
            workSheet.Cells[5, 2] = "Вопрос";
            workSheet.Cells[5, 3] = "Верный ответ";
            workSheet.Cells[5, 4] = "Ответ Пользователя";
            workSheet.Cells[5, 5] = "Балл";
            workSheet.Cells[5, 6] = "Длительность ответа";

            for (int i = 1; i <= qq; i++)
            {
                workSheet.Cells[5 + i, 1] = i;
                workSheet.Cells[5 + i, 2] = report[i-1, 0];
                workSheet.Cells[5 + i, 3] = report[i-1, 1];
                workSheet.Cells[5 + i, 4] = report[i-1, 2];
                workSheet.Cells[5 + i, 5] = report[i-1, 3];
                workSheet.Cells[5 + i, 6] = report[i-1, 4];
                score += Convert.ToInt32(report[i-1, 3]);
            }

            workSheet.Cells[qq + 7, 1] = $"Пользователь набрал {score} из {qq} баллов";
            workSheet.Range[workSheet.Cells[qq + 7, 1], workSheet.Cells[qq + 7, 8]].Merge();

            workSheet.Cells.Font.Size = 10;
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
            workSheet.Columns[3].AutoFit();
            workSheet.Columns[4].AutoFit();
            workSheet.Columns[5].AutoFit();
            workSheet.Columns[6].AutoFit();

            string currentTime = DateTime.Now.ToString().Replace(':', '_');
            string fileName = ($"{Environment.CurrentDirectory}\\Report\\{userName}-{currentTime}.xlsx");
            workBook.SaveAs(fileName);

            CloseExcel(xlRange, workBook, workSheet, excel);

 
        }

       

        public static void Questions(string[] questions, int countQuestions, int count, int qq, string[,] report)
        {
            string rightAnswer = questions[1];
            string answer = "";
            string point = "0";
            DateTime t1, t2;
            t1 = DateTime.Now;

            Shuffle(questions, count);
            Console.WriteLine();
            Console.WriteLine("Вопрос {0} из {1}: " + questions[0], countQuestions, qq);
            Console.WriteLine();
            for (int j = 1; j < count; j++)
            {
                Console.WriteLine("Вариант {0}: " + questions[j], j);
            }
            Console.WriteLine();
            Console.Write("Ваш вариант ответа: ");

            while (true)
            {
                string text = Console.ReadLine();

                if (int.TryParse(text, out int number))
                {
                    int key = Convert.ToInt32(text);
                    if (key > 0 && key < count)
                    {
                        answer = Convert.ToString(questions[key]);

                        break;
                    }
                    else
                    {
                        Console.Write("Такой вариант отсутствует, повторите ввод от 1 до {0}: ", (count - 1));
                    }
                }
                else
                {
                    Console.Write("Необходимо ввести цифру от 1 до {0}: ", (count - 1));
                }
            }
            t2 = DateTime.Now;
            if (answer == rightAnswer)
            {
                point = "1";
            }
            else
            {
                point = "0";
            }

            TimeSpan ts = t2 - t1;

            string time = Convert.ToString(ts.Hours.ToString() + " минут " + ts.Seconds.ToString() + " секунд");

            report[countQuestions - 1, 0] = questions[0]; //внесение вопроса
            report[countQuestions - 1, 1] = rightAnswer; //внесение верного ответа
            report[countQuestions - 1, 2] = answer; //внесение ответа Юзера
            report[countQuestions - 1, 3] = point; //внесение ответа Юзера
            report[countQuestions - 1, 4] = time; //внесение затраченного времени на ответ
            
        }
        public static void Shuffle(string[] arr, int count)
        {
            Random rand = new Random();

            for (int i = count - 1; i >= 1; i--)
            {
                int j = rand.Next(i);

                string tmp = arr[j + 1];
                arr[j + 1] = arr[i];
                arr[i] = tmp;
            }
        }
    }
}

