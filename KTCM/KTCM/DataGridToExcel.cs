using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Button = System.Windows.Forms.Button;
using DataTable = System.Data.DataTable;
using Rectangle = System.Drawing.Rectangle;
using System.Data.SQLite;

namespace KTCM
{
    internal class DataGridToExcel
    {
        static string connectionString = @"Data Source=ктсм.db;Version=3;";

        //Передача данных из грида в ексель DataGridViewToExcel
        public static void DataGridViewToExcel(DataGridView dataGridView)
        {
            /*//объявляем приложение
            Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
            //отобразить Excel
            ex.Visible = true;
            //количество листов в рабочей книге
            ex.SheetsInNewWorkbook = 1;
            //добавить рабочую книгу
            Excel.Workbook workbook = ex.Workbooks.Add(Type.Missing);
            //получаем первый лист документа
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
            //название листа вкладка снизу
            sheet.Name = "Отчет";
            for (int i = 1; i <= 9; i++)
            {
                for (int j = 1; j < 9; j++)
                    sheet.Cells[i, j] = String.Format("Boom {0} {1}", i, j);
            }
            //захватываем диапазон ячеек
            //Excel.Range range = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[9, 9]);
            sheet.Columns.EntireColumn.AutoFit();*/

            Microsoft.Office.Interop.Excel.Application excel = new Excel.Application(); //создаемCOM-объектExcel
            excel.Visible = true; //делаем объект видимым
            excel.SheetsInNewWorkbook = 1;//количество листов в книге
            excel.Workbooks.Add(Type.Missing); //добавляем книгу
            Excel.Workbook workbook = excel.Workbooks[1]; //получам ссылку на первую открытую книгу
            Excel.Worksheet sheet = (Worksheet)workbook.Worksheets.get_Item(1);//получаем ссылку на первый лист

            //рисуем линии вокруг шапки
            Excel.Range range = sheet.get_Range("F1", "F6");//выбираем ячейки
            //range.Merge(Type.Missing);
            //Устанавливаем стиль и толщину линии
            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.Weight = Excel.XlBorderWeight.xlHairline;

            //заголовок "Журнал"
            Excel.Range excelcells = sheet.get_Range("A1", "F3");//выбираем ячейки
            excelcells.Merge(Type.Missing);//объединяем ячейки
            //задаем выравнивание
            excelcells.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcells.VerticalAlignment = Excel.Constants.xlCenter;
            excelcells.Font.Size = 35;
            //excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;
            //excelcells.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            //excelcells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //excelcells.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcells.Value = "Журнал";//значение
            //следующая строка "уведомлений"
            Excel.Range excelcellsU = sheet.get_Range("A4", "F4");//выбираем ячейки
            excelcellsU.Merge(Type.Missing);//объединяем ячейки
            //задаем выравнивание
            excelcellsU.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcellsU.VerticalAlignment = Excel.Constants.xlCenter;
            excelcellsU.Font.Size = 15;
            //excelcellsU.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            //excelcellsU.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //excelcellsU.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcellsU.Value = "уведомлений";//значение
            //следующая строка диспетчера..
            Excel.Range excelcellsD = sheet.get_Range("A5", "F5");//выбираем ячейки
            excelcellsD.Merge(Type.Missing);//объединяем ячейки
            //задаем выравнивание
            excelcellsD.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcellsD.VerticalAlignment = Excel.Constants.xlCenter;
            excelcellsD.Font.Size = 13;
            //excelcellsD.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            //excelcellsD.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //excelcellsD.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcellsD.Value = "диспетчера по сигнализации и связи";//значение
            //следующая строка "при обслуживании...
            Excel.Range excelcellsP = sheet.get_Range("A6", "F6");//выбираем ячейки
            excelcellsP.Merge(Type.Missing);//объединяем ячейки
            //задаем выравнивание
            excelcellsP.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcellsP.VerticalAlignment = Excel.Constants.xlCenter;
            excelcellsP.Font.Size = 13;
            //excelcellsP.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            //excelcellsP.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //excelcellsP.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcellsP.Value = "при обслуживании устройств АСК ПС";//значение
            //следующая строка "КТСМ01Д"
            Excel.Range excelcellsK1 = sheet.get_Range("A7", "B7");//выбираем ячейки
            excelcellsK1.Merge(Type.Missing);//объединяем ячейки
            //задаем выравнивание
            excelcellsK1.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcellsK1.VerticalAlignment = Excel.Constants.xlCenter;
            excelcellsK1.Font.Size = 13;
            excelcellsK1.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            excelcellsK1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            excelcellsK1.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcellsK1.Value = "КТСМ01Д";//значение
            //следующая строка "СТП..
            Excel.Range excelcellsSTP1 = sheet.get_Range("C7", "F7");//выбираем ячейки
            excelcellsSTP1.Merge(Type.Missing);//объединяем ячейки
            //задаем выравнивание
            excelcellsSTP1.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcellsSTP1.VerticalAlignment = Excel.Constants.xlCenter;
            excelcellsSTP1.Font.Size = 13;
            excelcellsSTP1.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            excelcellsSTP1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            excelcellsSTP1.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcellsSTP1.Value = "СТП БЧ 19.346-2016 Пункт 7.5 ТК №25, ТК №27";//значение
            //следующая строка "КТСМ02
            Excel.Range excelcellsK2 = sheet.get_Range("A8", "B8");//выбираем ячейки
            excelcellsK2.Merge(Type.Missing);//объединяем ячейки
            //задаем выравнивание
            excelcellsK2.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcellsK2.VerticalAlignment = Excel.Constants.xlCenter;
            excelcellsK2.Font.Size = 13;
            excelcellsK2.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            excelcellsK2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            excelcellsK2.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcellsK2.Value = "КТСМ02";//значение
            //следующая строка "СТП..
            Excel.Range excelcellsSTP2 = sheet.get_Range("C8", "F8");//выбираем ячейки
            excelcellsSTP2.Merge(Type.Missing);//объединяем ячейки
            //задаем выравнивание
            excelcellsSTP2.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcellsSTP2.VerticalAlignment = Excel.Constants.xlCenter;
            excelcellsSTP2.Font.Size = 13;
            excelcellsSTP2.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            excelcellsSTP2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            excelcellsSTP2.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcellsSTP2.Value = "СТП БЧ 19.345-2017 Пункт 7.5";//значение
            //заголовок таблицы
            Excel.Range excelcellsS = sheet.get_Range("A9");//выбираем ячейки
            Excel.Range excelcellsSH = sheet.get_Range("B9");//выбираем ячейки
            Excel.Range excelcellsData = sheet.get_Range("C9");//выбираем ячейки
            Excel.Range excelcellsN = sheet.get_Range("D9");//выбираем ячейки
            Excel.Range excelcellsO = sheet.get_Range("E9");//выбираем ячейки
            Excel.Range excelcellsK = sheet.get_Range("F9");//выбираем ячейки
            //excelcellsZ.Merge(Type.Missing);//объединяем ячейки
            //задаем выравнивание
            excelcellsS.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcellsS.VerticalAlignment = Excel.Constants.xlCenter;
            excelcellsS.Font.Size = 13;
            excelcellsS.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            excelcellsS.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            excelcellsS.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcellsS.Value = "Станции";//значение
            //задаем выравнивание
            excelcellsSH.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcellsSH.VerticalAlignment = Excel.Constants.xlCenter;
            excelcellsSH.Font.Size = 13;
            excelcellsSH.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            excelcellsSH.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            excelcellsSH.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcellsSH.Value = "ШН";
            //задаем выравнивание
            excelcellsData.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcellsData.VerticalAlignment = Excel.Constants.xlCenter;
            excelcellsData.Font.Size = 13;
            excelcellsData.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            excelcellsData.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            excelcellsData.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcellsData.Value = "Дата";
            //задаем выравнивание
            excelcellsN.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcellsN.VerticalAlignment = Excel.Constants.xlCenter;
            excelcellsN.Font.Size = 13;
            excelcellsN.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            excelcellsN.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            excelcellsN.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcellsN.Value = "Начало работ";
            //задаем выравнивание
            excelcellsO.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcellsO.VerticalAlignment = Excel.Constants.xlCenter;
            excelcellsO.Font.Size = 13;
            excelcellsO.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            excelcellsO.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            excelcellsO.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcellsO.Value = "Окончание работ";
            //задаем выравнивание
            excelcellsK.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcellsK.VerticalAlignment = Excel.Constants.xlCenter;
            excelcellsK.Font.Size = 13;
            excelcellsK.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            excelcellsK.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            excelcellsK.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            excelcellsK.Value = "Тип аппаратуры";

            for (int j = 0; j < dataGridView.Columns.Count; ++j)
            {
                for (int i = 0; i < dataGridView.Rows.Count; ++i)
                {
                    Excel.Range cell = (Excel.Range)sheet.Cells[i + 10, j + 1];

                    object? Val = dataGridView.Rows[i].Cells[j].Value;
                    if (Val != null)
                        cell.Value2 = Val.ToString();

                    cell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    cell.Borders.Weight = Excel.XlBorderWeight.xlHairline;
                }
            }

            sheet.Columns.EntireColumn.AutoFit();
            sheet.Columns.AutoFit();
        }

        //void DataGridDataTime
        public static void DataGridDataTime(DataGridView dataGridView, string stringQuery)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                try
                {
                    connection.Open();
                }
                catch (SQLiteException ex)
                {
                    MessageBox.Show($"Нет подключения к базе данных {ex.Message}");
                    return;
                }
                using (SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter(stringQuery, connection))
                    try
                    {
                        DataSet dataSet = new DataSet();
                        dataAdapter.Fill(dataSet);
                        dataGridView.DataSource = dataSet.Tables[0];
                    }
                    catch (SQLiteException ex)
                    {
                        MessageBox.Show($"Ошибка {ex.Message}");
                    }
                    finally
                    {
                        connection.Close();
                    }
            }
        }
    }
}
