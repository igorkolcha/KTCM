using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
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
        public static void AddEmployee(TextBox textBox)
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
        }
        #endregion
    }

