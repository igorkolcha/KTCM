using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq.Expressions;
using System.Reflection.Emit;
using System.Text;
using System.Windows.Forms;

namespace KTCM
{
    internal class ConnectionDataBase
    {
        static string connectionString = @"Data Source=ктсм.db;Version=3;";

        #region method Connection
        public static void Connection(DataGridView dataGridView, string stringQuery)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                try
                {
                    connection.Open();
                }
                catch (SQLiteException ex)
                {
                    MessageBox.Show($"Ошибка подключения к базе данных SQLite: {ex.Message}");
                    return; // Выходим из метода, если нет подключения
                }
                using (SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter(stringQuery, connection))
                {
                    try
                    {
                        DataSet dataSet = new DataSet();
                        dataAdapter.Fill(dataSet);

                        // Настройка DataGridView
                        dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                        dataGridView.DataSource = dataSet.Tables[0];
                        dataGridView.RowHeadersVisible = false;
                        //dataGridView.Visible = true;
                        //dataGridView.AllowUserToAddRows = true;
                        dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                        //dataGridView1.Columns["фамилия"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                        // Обновить при изменении размера формы
                        dataGridView.Resize += (s, args) =>
                        {
                            dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
                        };
                    }
                    catch (SQLiteException ex)
                    {
                        // Обработка ошибок при выполнении запроса
                        MessageBox.Show($"Ошибка выполнения запроса SQLite: {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        // Общая обработка других ошибок
                        MessageBox.Show($"Произошла ошибка: {ex.Message}");
                    }
                    finally
                    {
                        // Соединение будет закрыто автоматически благодаря using,
                        // но явный вызов connection.Close() внутри finally, как у вас было,
                        // не нужен, но и не повредит, если вы решите убрать using.
                        // В данном случае 'using' гарантирует вызов Dispose(), который закроет соединение.
                        // connection.Close(); // Необязательно при использовании `using` для connection
                    }
                }
            }
        }
        #endregion

        #region method DeleteEmployee
        public static void DeleteEmployee(DataGridView dataGridView, System.Windows.Forms.Label label)
        {
            // Проверка входных данных
            if (label == null || string.IsNullOrWhiteSpace(label.Text))
            {
                MessageBox.Show("Не указано значение для удаления (метка пуста).");
                return;
            }

            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    // Параметризованный запрос — защищает от SQL-инъекций
                    string query = "DELETE FROM шн WHERE фамилия = @surname";
                    using (SQLiteCommand command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@surname", label.Text);

                        int rowsAffected = command.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            MessageBox.Show($"{rowsAffected}  запись удалена: {label.Text}");
                        }
                        else
                        {
                            MessageBox.Show($"Сотрудник с фамилией '{label.Text}' не найден.");
                        }
                    }
                }
                catch (SQLiteException ex)
                {
                    MessageBox.Show($"Ошибка базы данных: {ex.Message}", "Ошибка SQLite", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Неожиданная ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        #endregion

        #region method AddEmployee
        public static void AddEmployee(System.Windows.Forms.TextBox textBox)
        {
            if (textBox == null || string.IsNullOrWhiteSpace(textBox.Text))
            {
                MessageBox.Show("Не указано значение для записи.");
                return;
            }
            string insert = @"INSERT INTO шн (фамилия)
                             VALUES (@фамилия)";
            using var conn = new SQLiteConnection(connectionString);
            using var cmd = new SQLiteCommand(insert, conn);
            cmd.Parameters.AddWithValue("@фамилия", textBox.Text);
            conn.Open();
            cmd.ExecuteNonQuery();
        }
        #endregion

        #region method BeginWork
        public static void BeginWork(System.Windows.Forms.Button button, DateTimePicker dateTimePicker, DataGridView dataGridView)
        {
            string[] ktsmArray = { "лучеса", "чепино", "гродок чет", "городок неч" };

            string ktsm;

            if (Array.IndexOf(ktsmArray, button.Text) >= 0)
            {
                ktsm = "ктсм1д";
            }
            else
                ktsm = "ктсм2";

            /*string insert = @"INSERT INTO шн (дата, станции, фамилия, начало, месяц, ктсм) 
                               VALUES (@дата, @станции, @фамилия, @начало, @месяц, @ктсм)";
            using var conn = new SQLiteConnection(connectionString);
            using var cmd = new SQLiteCommand(insert, conn);
            cmd.Parameters.Add("@дата", DbType.String).Value = dateTimePicker.Value.ToShortDateString();
            cmd.Parameters.Add("@станции", DbType.String).Value = button.Text;
            cmd.Parameters.Add("@фамилия", DbType.String).Value = dataGridView.CurrentCell?.Value?.ToString();
            cmd.Parameters.Add("@начало", DbType.String).Value = DateTime.Now.ToShortTimeString();
            cmd.Parameters.Add("@месяц", DbType.Int32).Value = dateTimePicker.Value.Month;
            cmd.Parameters.Add("@ктсм", DbType.String).Value = ktsm;
            conn.Open();
            cmd.ExecuteNonQuery();

            int countENQ = cmd.ExecuteNonQuery();
            //int count = dataAdapter.Update(dataTable);
            MessageBox.Show("Начало работ на КТСМ в " + DateTime.Now.ToShortTimeString(),
                "Изменено записей: " + countENQ);*/
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                try
                {
                    connection.Open();
                }
                catch (SQLiteException ex)
                {
                    MessageBox.Show($"Нет подключения к базе данных {ex.Message}");
                }
                using (SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter())
                {
                    try
                    {
                        //DataTable dataTable = new DataTable("ктсм");
                        SQLiteCommand command = connection.CreateCommand();
                        command.CommandText = "INSERT INTO ктсм (дата, станции, фамилия, начало, месяц, ктсм) " +
                            "VALUES ('" + dateTimePicker.Value.ToShortDateString() + "','" + button.Text + "','" +
                            dataGridView.CurrentCell?.Value?.ToString() + "" + "','" + DateTime.Now.ToShortTimeString() + "','" +
                            dateTimePicker.Value.Month + "', '" + ktsm + "')";

                        int countENQ = command.ExecuteNonQuery();
                        //int count = dataAdapter.Update(dataTable);
                        MessageBox.Show("Начало работ на КТСМ в " + DateTime.Now.ToShortTimeString(),
                            "Изменено записей: " + countENQ);
                    }
                    catch (SQLiteException ex)
                    {
                        MessageBox.Show($"Error: {ex.Message}");
                    }
                    connection.Close();
                }
            }
        }
        #endregion

        #region method EndWork
        public static void EndWork(System.Windows.Forms.Button button, DateTimePicker dateTimePicker, string  stringQuery)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                try
                {
                    connection.Open();
                }
                catch (SQLiteException ex) { MessageBox.Show($"Нет подключения к базе данных {ex.Message}"); }
                using (SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter(stringQuery, connection))
                {
                    try
                    {
                        //DataTable dataTable = new DataTable("шн");
                        SQLiteCommand command = connection.CreateCommand();
                        command.CommandText = "UPDATE ктсм SET конец = '" + DateTime.Now.ToShortTimeString() + "'" +
                            "WHERE станции ='" + button.Text + "' AND дата = '" + dateTimePicker.Value.ToShortDateString() + "'";

                        int countEND = command.ExecuteNonQuery();
                        MessageBox.Show("на ктсм работа закончена в " + DateTime.Now.ToShortTimeString(), "изменено записей: " + countEND);
                    }
                    catch (SQLiteException ex) { MessageBox.Show($"Error: {ex.Message}"); }
                }
            } 
        }
        #endregion
    }
}

