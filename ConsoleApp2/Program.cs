using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;





namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            
            

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open("D:\\Questions.xlsx");

            //int colCount = xlRange.Columns.Count; //Колонки
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
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[level];
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
            //Report(userName, level, report, questionsQuontyty);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //Marshal.ReleaseComObject(xlRange);
            //Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            // Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();

            //Marshal.ReleaseComObject(xlApp);
            ReportToExcel(userName, level, report, questionsQuontyty);
            Console.CursorVisible = false;
            Console.WriteLine("Данные сохранены.\nНажмите <Enter> для завершения программы");
            Console.Read();
        }





        static void ReportToExcel(string userName, int level, string[,] report, int qq)
        {
            int score = 0;
            Excel.Application excel;
            Excel.Workbook worKbooK;
            Excel.Worksheet worKsheeT;
            //Excel.Range celLrangE;


            excel = new Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            worKbooK = excel.Workbooks.Add(Type.Missing);


            worKsheeT = (Excel.Worksheet)worKbooK.ActiveSheet;
            worKsheeT.Name = userName;

            worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, 8]].Merge();
            worKsheeT.Range[worKsheeT.Cells[2, 1], worKsheeT.Cells[2, 8]].Merge();
            worKsheeT.Range[worKsheeT.Cells[3, 1], worKsheeT.Cells[3, 8]].Merge();
            worKsheeT.Cells[1, 1] = $"Результаты тестирования пользователя {userName}";
            worKsheeT.Cells[2, 1] = $"Время тестирования: {DateTime.Now.ToString()}";
            worKsheeT.Cells[3, 1] = $"Уровень сложности: {level}";
            worKsheeT.Cells[5, 1] = "№";
            worKsheeT.Cells[5, 2] = "Вопрос";
            worKsheeT.Cells[5, 3] = "Верный ответ";
            worKsheeT.Cells[5, 4] = "Ответ Пользователя";
            worKsheeT.Cells[5, 5] = "Балл";
            worKsheeT.Cells[5, 6] = "Длительность ответа";



            for (int i = 1; i <= qq; i++)
            {
                worKsheeT.Cells[5 + i, 1] = i;
                worKsheeT.Cells[5 + i, 2] = report[i-1, 0];
                worKsheeT.Cells[5 + i, 3] = report[i-1, 1];
                worKsheeT.Cells[5 + i, 4] = report[i-1, 2];
                worKsheeT.Cells[5 + i, 5] = report[i-1, 3];
                worKsheeT.Cells[5 + i, 6] = report[i-1, 4];
                score += Convert.ToInt32(report[i-1, 3]);
            }

            worKsheeT.Cells[qq + 7, 1] = $"Пользователь набрал {score} из {qq} баллов";
            worKsheeT.Range[worKsheeT.Cells[qq + 7, 1], worKsheeT.Cells[qq + 7, 8]].Merge();

            worKsheeT.Cells.Font.Size = 10;
            worKsheeT.Columns[1].AutoFit();
            worKsheeT.Columns[2].AutoFit();
            worKsheeT.Columns[3].AutoFit();
            worKsheeT.Columns[4].AutoFit();
            worKsheeT.Columns[5].AutoFit();
            worKsheeT.Columns[6].AutoFit();

            string fileName = ($"\\Report\\{userName}-{DateTime.Now.ToString()}.xlsx");
            string fileName2 = "D:" + fileName.Replace(':', '_');
            //Console.WriteLine(fileName2);


            worKbooK.SaveAs(fileName2);
            
            worKbooK.Close();
            excel.Quit();


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
                //Console.WriteLine("Ответ {0} - верный.", answer);
                point = "1";
            }
            else
            {
                //Console.WriteLine("Ответ {0} - не верный.\nПрвельный ответ: {1}.",answer, rightAnswer);
                point = "0";
            }

            TimeSpan ts = t2 - t1;

            string time = Convert.ToString(ts.Hours.ToString() + " минут " + ts.Seconds.ToString() + " секунд");
            //Console.WriteLine("Время ответа: " + ts.Hours.ToString() + " минут " + ts.Seconds.ToString()+" секунд");
            //Console.Read();

            report[countQuestions - 1, 0] = questions[0]; //внесение вопроса
            report[countQuestions - 1, 1] = rightAnswer; //внесение верного ответа
            report[countQuestions - 1, 2] = answer; //внесение ответа Юзера
            report[countQuestions - 1, 3] = point; //внесение ответа Юзера
            report[countQuestions - 1, 4] = time; //внесение затраченного времени на ответ
            //Console.WriteLine();
            //Console.WriteLine("Нажмите <Enter> клавишу для перехода к следующему вопросу");
            //Console.Read();
            
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

