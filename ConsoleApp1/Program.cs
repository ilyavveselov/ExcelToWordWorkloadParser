using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
{
    Console.WriteLine("Привет! Я переношу данные из excel-файла в word. Для успешного переноса ваш excel-файл должен соответствовать следующим условиям:\n" +
        "C столбец - семестр, D - Дисциплина, E - специальность, F - курс, S-ЛЕК, T - ЛАБ, U - ПРА, V - ЭКЗ, W - ЗАЧ, X - К.ПР, Y - К.Р, AA - КОНС");

    try
    {
        Microsoft.Office.Interop.Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "GetWindowThreadProcessId")]
        static extern int GetWindowThreadProcessId([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd, out int lpdwProcessId);
        void KillExcel()
        {
            int ExcelPID = 0;
            int Hwnd = 0;
            Hwnd = ex.Hwnd;
            System.Diagnostics.Process ExcelProcess;
            GetWindowThreadProcessId((IntPtr)Hwnd, out ExcelPID);
            ExcelProcess = System.Diagnostics.Process.GetProcessById(ExcelPID);
            ////Конец подготовки к убийству процесса Excel

            ex.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ex);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            ////Убийство процесса Excel
            ExcelProcess.Kill();
            ExcelProcess = null;
        }
        ex.Visible = false;
        string pathToExcelFile = @Path.GetDirectoryName(Environment.CurrentDirectory) + "\\Нагрузка.xlsx";
        Excel.Workbook excelAppWorkbooks = null;
        try
        {
            excelAppWorkbooks = ex.Workbooks.Open(pathToExcelFile,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing);
        }
        catch
        {
            KillExcel();
            throw new InvalidOperationException("Файла EXCEL нет, либо у него неверное название. Верное - Нагрузка");
        }

        var excelSheets = excelAppWorkbooks.Worksheets;
        var excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(1);
        int counterOfNotEmptyCells = 0;
        string[] labels = new string[] { null, "дисциплина", "специальность", "курс", "ЛЕК", "ЛАБ", "ПРА", "ЭКЗ", "ЗАЧ", "К.ПР", "К.Р", "КОНС" };
        Console.WriteLine("Извлекаю данные из Excel файла...");
        int labelCounter = 0;
        string autumnLiteral = "о";
        string springLiteral = "в";
        List<string> AddInfo(string cell)
        {
            List<string> values = new List<string>();
            bool isCellNotNull = true;
            int counter = 2;
            var value = String.Empty;
            Excel.Range excelCells = excelWorksheet.get_Range($"{cell}1", Type.Missing); ;
            var labelOfCell = excelCells.Value2;
            if (excelCells.Value2 == labels[labelCounter])
            {
                while (isCellNotNull)
                {
                    string fullCell = $"{cell}{counter}";
                    excelCells = excelWorksheet.get_Range(fullCell, Type.Missing);
                    try
                    {
                        value = excelCells.Value2.ToString();
                        values.Add(value);
                        counter++;
                        counterOfNotEmptyCells++;
                    }
                    catch
                    {
                        isCellNotNull = false;
                    }
                }
                if (counterOfNotEmptyCells == 0)
                {
                    excelAppWorkbooks.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelSheets);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppWorkbooks);
                    KillExcel();
                    throw new InvalidOperationException("В EXCEL файле нет столбцов, либо они расположены не по шаблону");
                }
                Console.WriteLine($"\tСтолбец {labels[labelCounter]} извлечен!");
                labelCounter++;
                return values;
            }
            else
            {
                excelAppWorkbooks.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelSheets);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppWorkbooks);
                KillExcel();
                throw new InvalidOperationException($"Вместо столбца {labelOfCell} должен был быть {labels[labelCounter]}. Excel файл не подходит под шаблон. Выхожу...");
            }
        }
        List<double> AddInfoNumbers(string cell)
        {
            int counterOfNull = 0;
            List<double> values = new List<double>();
            int counter = 2;
            double value = 0;
            Excel.Range excelCells = excelWorksheet.get_Range($"{cell}1", Type.Missing);
            var labelOfCell = excelCells.Value2;
            if (excelCells.Value2 == labels[labelCounter])
            {
                for (int i = 0; i < counterOfNotEmptyCells; i++)
                {
                    string fullCell = $"{cell}{counter}";
                    excelCells = excelWorksheet.get_Range(fullCell, Type.Missing);
                    if (excelCells.Value2 == null)
                    {
                        value = 0;
                        values.Add(value);
                        counter++;
                        counterOfNull++;
                    }
                    else if (excelCells.MergeArea.CountLarge > 1)
                    {
                        value = Math.Round((excelCells.Value2), MidpointRounding.AwayFromZero);
                        var len = excelCells.MergeArea.CountLarge;
                        for (int j = 0; j < len; j++)
                        {
                            if (j == 0)
                                values.Add(value);
                            else
                                values.Add(0);
                            i++;
                            counter++;
                        }
                        i--;
                    }
                    else
                    {
                        value = Math.Round((excelCells.Value2), MidpointRounding.AwayFromZero);
                        values.Add(value);
                        counter++;
                    }
                }
                if (counterOfNull == counterOfNotEmptyCells)
                {
                    excelAppWorkbooks.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelSheets);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppWorkbooks);
                    KillExcel();
                    throw new InvalidOperationException($"Столбец {labels[labelCounter]} оказался пустым. Выхожу...");
                }
                Console.WriteLine($"\tСтолбец {labels[labelCounter]} извлечен!");
                labelCounter++;
                return values;
            }
            else
            {
                excelAppWorkbooks.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelSheets);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppWorkbooks);
                KillExcel();
                throw new InvalidOperationException($"Вместо столбца {labelOfCell} должен был быть {labels[labelCounter]}. Excel файл не подходит под шаблон. Выхожу...");
            }
        }
        List<string> GroupsFromNumbersToLetters(List<string> groupsNum)
        {
            List<string> values = new List<string>();
            foreach (var item in groupsNum)
            {
                switch (item.Trim())
                {
                    case "09.03.01":
                        {
                            values.Add("ИСТ");
                            break;
                        }
                    case "09.03.02":
                        {
                            values.Add("ВТ");
                            break;
                        }
                    case "09.05.01":
                        {
                            values.Add("АС");
                            break;
                        }
                    default:
                        break;
                }
            }
            return values;
        }
        var semester = AddInfo("C");
        if (semester.Contains(autumnLiteral) == false)
        {
            excelAppWorkbooks.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelSheets);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppWorkbooks);
            KillExcel();
            throw new InvalidOperationException("Первого семестра нет в нагрузке, либо столбец с семестравми вовсе отсутствует");
        }
        List<int> whereIsSpring = new List<int>();
        for (int i = 0; i < semester.Count; i++)
        {
            if (semester[i] != autumnLiteral)
            {
                whereIsSpring.Add(i);
            }
        }
        if (whereIsSpring.Count == 0)
        {
            excelAppWorkbooks.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelSheets);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppWorkbooks);
            KillExcel();
            throw new InvalidOperationException("Второго семестра нет в нагрузке, либо столбец с семестравми вовсе отсутствует");
        }

        var subjectNames = AddInfo("D");
        var groupsNumbers = AddInfo("E");
        counterOfNotEmptyCells /= 3;
        var groupsCourse = AddInfoNumbers("F");
        var lectures = AddInfoNumbers("S");
        var lab = AddInfoNumbers("T");
        var practice = AddInfoNumbers("U");
        var exam = AddInfoNumbers("V");
        var credit = AddInfoNumbers("W");
        var courseWork = AddInfoNumbers("X");
        var vkr = AddInfoNumbers("Y");
        var consult = AddInfoNumbers("AA");
        excelAppWorkbooks.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelSheets);
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppWorkbooks);
        KillExcel();

        var groups = GroupsFromNumbersToLetters(groupsNumbers);
        List<string> subjectsIncapsulated = new List<string>(subjectNames);
        List<string> groupsIncapsulated = new List<string>(groups);
        List<string> unique = new List<string>();

        Console.WriteLine("Формирую строки для вывода в Word...");

        foreach (var item in subjectNames)
        {
            if (unique.Contains(item.Trim()) == false)
            {
                unique.Add(item.Trim());
            }
        }
        List<string> shapedRows = new List<string>();
        string MakeResult(string subjectName)
        {
            double lectHA = 0;
            double labHA = 0;
            double practtHA = 0;
            double examHA = 0;
            double creditHA = 0;
            double cwHA = 0;
            double vkrHA = 0;
            double consultHA = 0;
            double lectHS = 0;
            double labHS = 0;
            double practtHS = 0;
            double examHS = 0;
            double creditHS = 0;
            double cwHS = 0;
            double vkrHS = 0;
            double consultHS = 0;
            List<string> groupsInResult = new List<string>();
            List<double> coursesInResult = new List<double>();

            for (int i = 0; i < subjectNames.Count; i++)
            {
                if (subjectNames[i] == subjectName)
                {
                    if (semester[i] == autumnLiteral)
                    {
                        lectHA += (lectures[i]);
                        labHA += (lab[i]);
                        practtHA += (practice[i]);
                        examHA += (exam[i]);
                        creditHA += (credit[i]);
                        cwHA += (courseWork[i]);
                        vkrHA += (vkr[i]);
                        consultHA += (consult[i]);
                    }
                    if (semester[i] == springLiteral)
                    {
                        lectHS += (lectures[i]);
                        labHS += (lab[i]);
                        practtHS += (practice[i]);
                        examHS += (exam[i]);
                        creditHS += (credit[i]);
                        cwHS += (courseWork[i]);
                        vkrHS += (vkr[i]);
                        consultHS += (consult[i]);
                    }
                    groupsInResult.Add(groups[i]);
                    coursesInResult.Add(groupsCourse[i]);
                    subjectNames[i] = "";
                }
            }
            List<string> groupsCoursesResult = new List<string>();
            for (int i = 0; i < groupsInResult.Count; i++)
            {
                groupsCoursesResult.Add(groupsInResult[i] + "- 1" + coursesInResult[i].ToString());
            }
            var result = ($" {string.Join(',', groupsCoursesResult)} {subjectName} :" +
                $"{lectHA};{practtHA};{labHA};{consultHA};{creditHA};{examHA};{cwHA}:" +
                $"{lectHS};{practtHS};{labHS};{consultHS};{creditHS};{examHS};{cwHS}");
            return result;
        }
        foreach (var item in unique)
        {
            shapedRows.Add(MakeResult(item));
        }

        Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
        word.Visible = false;
        object miss = System.Reflection.Missing.Value;
        object path = @Path.GetDirectoryName(Environment.CurrentDirectory) + "\\Template\\ИП XXX";
        object pathToSave = @Path.GetDirectoryName(Environment.CurrentDirectory) + "\\ИП XXX";
        object readOnly = false;
        Document doc;
        Console.WriteLine("Открываю Word файл для вставки...");
        try
        {
            doc = word.Documents.Open
             (ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
        }
        catch
        {
            word.Quit();
            throw new InvalidOperationException("Файла с шаблоном WORD нет, либо он назван не по шаблону. Верно - ИП XXX. Формат doc");
        }
        try
        {
            word.Application.ActiveDocument.SaveAs(ref pathToSave, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss); ;
            doc.Close();
            doc = word.Documents.Open
             (ref pathToSave, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
        }
        catch (Exception exc)
        {
            doc.Close();
            word.Quit();
            throw new InvalidOperationException(exc.Message.ToString());
        }

        //num: 0-lectures 1-lab 2-practice 3-exam 4-credit 5-courseWork 6-consult 7-vkr
        double[] hoursResult = new double[8];
        double[] hoursFS = new double[8];
        double[] hoursSS = new double[8];

        Word.Range changeStyle(Word.Range wordCell)
        {
            wordCell.Bold = 0;
            wordCell.HighlightColorIndex = 0;
            wordCell.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            return wordCell;
        }
        int AddToWord(Table table, Microsoft.Office.Interop.Word.Range wordCell, int num, int counterForCell)
        {
            int counter = counterForCell;
            int skipedRowsCounter = 0;
            for (int i = counter; i < shapedRows.Count + counter; i++)
            {
                var splitedResult = shapedRows[i - counter].Split(':');
                var hour = double.Parse(splitedResult[1].Split(';')[num]);
                int column = 3;
                bool isAdded = false;
                if (hour != 0)
                {
                    changeStyle(wordCell);
                    wordCell.Font.Size = 10;
                    if (table.Cell(i - skipedRowsCounter, 2).Range.Text.Trim() == "Итого\r\a")
                    {
                        wordCell = table.Cell(i - skipedRowsCounter - 1, 2).Range;
                        wordCell.Rows.Add();
                        isAdded = true;
                    }
                    wordCell = table.Cell(i - skipedRowsCounter, 2).Range;
                    changeStyle(wordCell);
                    wordCell.Font.Size = 10;
                    wordCell.Text = $"{splitedResult[0]}\r\a";
                    wordCell = table.Cell(i - skipedRowsCounter, column).Range;
                    changeStyle(wordCell);
                    wordCell.Font.Size = 10;
                    wordCell.Text = hour.ToString();
                    counterForCell++;
                    hoursResult[num] += hour;
                    hoursFS[num] += hour;
                }
                else
                {
                    skipedRowsCounter++;
                }
                if (splitedResult[2].Split(';')[num] != "0")
                {
                    if (!isAdded)
                    {
                        changeStyle(wordCell);
                        wordCell.Font.Size = 10;
                        skipedRowsCounter--;
                        if (table.Cell(i - skipedRowsCounter, 2).Range.Text.Trim() == "Итого\r\a")
                        {
                            wordCell = table.Cell(i - skipedRowsCounter - 1, 2).Range;
                            wordCell.Rows.Add();
                        }
                        wordCell = table.Cell(i - skipedRowsCounter, 2).Range;
                        changeStyle(wordCell);
                        wordCell.Font.Size = 10;
                        wordCell.Text = $"{splitedResult[0]}\r\a";
                        counterForCell++;
                    }
                    hour = double.Parse(splitedResult[2].Split(';')[num]);
                    column = 5;
                    wordCell = table.Cell(i - skipedRowsCounter, column).Range;
                    wordCell.Text = hour.ToString();
                    hoursResult[num] += hour;
                    hoursSS[num] += hour;
                }
            }
            changeStyle(wordCell);
            wordCell = table.Cell(counterForCell, 3).Range;
            wordCell.Text = hoursResult[num].ToString();
            return counterForCell;
        }
        int counterForCell = 6;
        Table table;
        Word.Range wordCell;
        try
        {
            table = doc.Tables[2];
            wordCell = table.Cell(5, 2).Range;
        }
        catch
        {
            doc.Close();
            word.Quit();
            throw new InvalidOperationException(
                "В Word файле нет таблицы по Учебной работе");
        }
        Console.WriteLine("Вставляю данные в Учебную работу...");
        try
        {
            for (int ii = 0; ii < 7; ii++)
            {
                counterForCell = AddToWord(table, wordCell, ii, counterForCell) + 2;
                wordCell = table.Cell(counterForCell - 1, 2).Range;
                if (wordCell.Text.Trim() == "Контроль:\r\a")
                {
                    counterForCell += 1;
                }
                wordCell = table.Cell(counterForCell, 2).Range;
            }
            int counterForVrk = 0;
            foreach (var item in vkr)
            {
                if (!whereIsSpring.Contains(counterForVrk))
                {
                    hoursFS[7] += item;
                }
                else
                {
                    hoursSS[7] += item;
                }
                hoursResult[7] += item;
                counterForVrk++;
            }
        }
        catch
        {
            doc.Close();
            word.Quit();
            throw new InvalidOperationException("В WORD файле места для вставки в Учебную работу расположены не по шаблону, либо вовсе отсутствуют");
        }
        Console.WriteLine("Вставляю данные в Общие часы...");
        void AddToCommonTable(int row, double[] hours)
        {
            table = doc.Tables[7];
            int counterForSum = 3;
            wordCell = table.Cell(row, counterForSum).Range;
            double sumInHoursResult = 0;
            double[] tempHours = new double[hours.Length];
            for (int i = 0; i < hours.Length; i++)
            {
                tempHours[i] = hours[i];
            }
            tempHours[3] = hours[4];
            tempHours[4] = hours[5];
            tempHours[5] = hours[3];
            foreach (var item in tempHours)
            {
                changeStyle(wordCell);
                wordCell.Text = item.ToString();
                counterForSum++;
                wordCell = table.Cell(row, counterForSum).Range;
                sumInHoursResult += item;
            }
            counterForSum += 3;
            wordCell = table.Cell(row, counterForSum).Range;
            changeStyle(wordCell);
            wordCell.Text = sumInHoursResult.ToString();
        }
        try
        {
            AddToCommonTable(2, hoursFS);
            AddToCommonTable(4, hoursSS);
            AddToCommonTable(6, hoursResult);
        }
        catch (Exception)
        {
            doc.Close();
            word.Quit();
            throw new InvalidOperationException("В WORD файле места для вставки в Общую нагрузку расположены не по шаблону, либо вовсе отсутствуют");
        }
        try
        {
            word.Application.ActiveDocument.SaveAs(pathToSave);
        }
        catch (Exception)
        {
            doc.Close();
            word.Quit();
            throw new InvalidOperationException("Не получилось сохранить, так как папка, где лежал документ, была удалена. Попробуйте сохранить в другое место");
        }

        doc.Close();
        word.Quit();
        Console.WriteLine("Операция прошла успешно!");
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
        Thread.Sleep(7000);
    }
}