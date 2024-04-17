using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.ComponentModel;

namespace DataBaseWorker
{
    /// <summary>
    /// Класс работы с базой данных
    /// </summary>
    public class DataBaseApplication
    {
        private const string BaseConnection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
        /// <summary>
        /// Строка подключения
        /// </summary>
        public string ConnectionString { get; set; }

        private OleDbConnection conection = new OleDbConnection();
        private OleDbDataAdapter waitoperation = new OleDbDataAdapter();
        private DataSet celloperation = new DataSet();
        /// <summary>
        /// Возращает соединение с базой данных
        /// </summary>
        public OleDbConnection Connection { get => conection; }

        /// <summary>
        /// Конструктор по умолчанию
        /// </summary>
        public DataBaseApplication() { }

        /// <summary>
        /// Метод открытия соединения
        /// </summary>
        /// <returns>true если успешно</returns>
        public bool TryOpen()
        {
            conection.ConnectionString = ConnectionString;
            try
            {
                conection.Open();
            }
            catch(Exception e)
            {
                MessageBox.Show("Ошибка соединения:\n" + e.Message, "Система", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }
        /// <summary>
        /// Метод открытия соединения
        /// </summary>
        /// <param name="NameBase"></param>
        /// <returns>true если успешно</returns>
        public bool TryOpen(string NameBase)
        {
            ConnectionString = BaseConnection + NameBase;
            return TryOpen();
        }

        /// <summary>
        /// Модальное окно конструктора соединения
        /// </summary>
        public void BuildConnection()
        {
            Form form = new Form();
            form.FormBorderStyle = FormBorderStyle.None;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.Size = new Size(500, 300);

            TextBox DataBaseName1 = new TextBox();
            ComboBox ProviderName1 = new ComboBox();
            Button DialogResultOk1 = new Button();
            Button DialogResultCancel1 = new Button();
            OpenFileDialog openFile = new OpenFileDialog();
            Button MoreFile1 = new Button()
            {
                Text = "...",
                AutoSize = true,
                Width = 30
            };

            DataBaseName1.Size = new Size(form.Width / 3 * 2 - (MoreFile1.Width + 3), 20);
            DataBaseName1.Location = new Point((form.Width - (DataBaseName1.Width + MoreFile1.Width + 3)) / 2, form.Height / 3);
            Label DataBaseName2 = new Label()
            {
                Text = "Название базы данных и расширение",
                Location = new Point(DataBaseName1.Left, DataBaseName1.Top),
                AutoSize = true,
                Width = DataBaseName1.Width,
                Height = 20
            };
            DataBaseName2.Top -= DataBaseName2.Height;

            MoreFile1.Height = DataBaseName1.Height;
            MoreFile1.Left = DataBaseName1.Left + DataBaseName1.Width + 3;
            MoreFile1.Top = DataBaseName1.Top - 2;
            MoreFile1.Click += (sender, e) =>
            {
                if(openFile.ShowDialog() == DialogResult.OK)
                    DataBaseName1.Text = openFile.SafeFileName;
            };

            ProviderName1.Size = new Size(form.Width / 3 * 2, 20);
            ProviderName1.Location = new Point((form.Width - ProviderName1.Width) / 2, DataBaseName1.Top + DataBaseName1.Height + ProviderName1.Height * 2);
            ProviderName1.Items.Add("Microsoft.ACE.OLEDB.12.0");
            ProviderName1.SelectedIndex = 0;
            Label ProviderName2 = new Label()
            {
                Text = "Провайдер",
                Location = new Point(ProviderName1.Left, ProviderName1.Top),
                AutoSize = true,
                Width = ProviderName1.Width,
                Height = 20
            };
            ProviderName2.Top -= ProviderName2.Height;

            DialogResultOk1.Size = new Size(100, 50);
            DialogResultOk1.Location = new Point(ProviderName1.Left, form.Height / 4 * 3);
            DialogResultOk1.Text = "ОК";
            DialogResultOk1.Click += (sender, e) => form.DialogResult = DialogResult.OK;

            DialogResultCancel1.Size = new Size(100, 50);
            DialogResultCancel1.Location = new Point(ProviderName1.Left + ProviderName1.Width - DialogResultCancel1.Width, form.Height / 4 * 3);
            DialogResultCancel1.Text = "Отмена";
            DialogResultCancel1.Click += (sender, e) => form.DialogResult = DialogResult.Cancel;

            form.Controls.Add(DataBaseName1);
            form.Controls.Add(ProviderName1);
            form.Controls.Add(DialogResultOk1);
            form.Controls.Add(DialogResultCancel1);
            form.Controls.Add(DataBaseName2);
            form.Controls.Add(ProviderName2);
            form.Controls.Add(MoreFile1);

            if(form.ShowDialog() == DialogResult.OK)
                ConnectionString = string.Format("Provider={0};Data Source={1}", ProviderName1.SelectedItem, DataBaseName1.Text);
        }

        /// <summary>
        /// Закрывает соединение
        /// </summary>
        /// <returns> true если успешно</returns>
        public bool TryClose()
        {
            if(Connection.State == System.Data.ConnectionState.Open)
            {
                Connection.Close();
                return true;
            }

            return false;
        }

        /// <summary>
        /// Совращает таблицу из базы данных
        /// </summary>
        /// <param name="NameTable">Имя таблицы</param>
        /// <returns>Таблица из Базы данных</returns>
        public DataTable GetDataTable(string NameTable)
        {
            bool closed = Connection.State == ConnectionState.Closed;
            if (closed) TryOpen();

            OleDbDataAdapter adapter = new OleDbDataAdapter($"SELECT * FROM {NameTable}", Connection);
            DataSet data = new DataSet();

            adapter.Fill(data);

            if (closed) TryClose();

            return data.Tables[0];
        }
        /// <summary>
        /// Добавляет строку в базу данных
        /// </summary>
        /// <param name="NameTable">Название таблицы</param>
        /// <param name="values">Значения</param>
        public void AddDataRow(string NameTable, params object[] values)
        {
            bool closed = Connection.State == ConnectionState.Closed;
            if (closed) TryOpen();

            OleDbDataAdapter adapter = new OleDbDataAdapter($"SELECT * FROM {NameTable}", Connection);
            DataSet data = new DataSet();
            DataTable table = null;

            try
            {
                adapter.Fill(data);
                table = data.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка операции: " + ex.Message, "Система", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (closed) TryClose();
                return;
            }

            DataRow row = table.NewRow();

            try
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    if (i < values.Length)
                        row[i] = values[i];
                    else
                        row[i] = null;
                }

                table.Rows.Add(row);
                OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);
                adapter.Update(data);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка заполнения: " + ex.Message, "Система", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (closed) TryClose();
            }
        }
        /// <summary>
        /// Удаляет строку из базы данных по параметру
        /// </summary>
        /// <param name="NameTable">Название таблицы</param>
        /// <param name="parametr">Атрибут поиска</param>
        public void RemoveDataRow(string NameTable, object parametr)
        {
            bool closed = Connection.State == ConnectionState.Closed;
            if (closed) TryOpen();

            OleDbDataAdapter adapter = new OleDbDataAdapter($"SELECT * FROM {NameTable}", Connection);
            DataSet data = new DataSet();
            DataTable table = null;

            bool equal = false;

            try
            {
                adapter.Fill(data);
                table = data.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка операции: " + ex.Message, "Система", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (closed) TryClose();
                return;
            }

            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    equal = table.Rows[i][j].Equals(parametr);
                    if (!equal) continue;
                    table.Rows[i].Delete();
                    break;
                }
                if (equal) break;
            }


            try
            {
                OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);
                adapter.Update(data);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Ошибка удаления: " + ex.Message, "Система", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (closed) TryClose();
            }
        }
        /// <summary>
        /// Удаляет строку из базы данных по коду
        /// </summary>
        /// <param name="NameTable">Название базы данных</param>
        /// <param name="id">Код строки</param>
        public void RemoveAtDataRow(string NameTable, int id)
        {
            bool closed = Connection.State == ConnectionState.Closed;
            if (closed) TryOpen();

            OleDbDataAdapter adapter = new OleDbDataAdapter($"SELECT * FROM {NameTable}", Connection);
            DataSet data = new DataSet();
            DataTable table = null;

            try
            {
                adapter.Fill(data);
                table = data.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка операции: " + ex.Message, "Система", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (closed) TryClose();
                return;
            }

            int low = 0;
            int higth = table.Rows.Count - 1;
            int mid = 0;
            int value = 0;

            while (low <= higth)
            {
                mid = (low + higth) / 2;
                value = (int)table.Rows[mid][0];

                if (value == id)
                {
                    table.Rows[mid].Delete();
                    break;
                }

                if (value < id)
                    low = mid + 1;
                else
                    higth = mid - 1;
            }

            try
            {
                OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);
                adapter.Update(data);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка удаления: " + ex.Message, "Система", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (closed) TryClose();
            }
        }
        /// <summary>
        /// Открывает поток редактирования
        /// </summary>
        /// <param name="NameTable">Название таблицы</param>
        /// <param name="id">Код строки</param>
        /// <param name="row">Возращаемая строка</param>
        public void BeginUpdateDataRowAt(string NameTable, int id, ref DataRow row)
        {
            if (Connection.State == ConnectionState.Closed) TryOpen();

            waitoperation.SelectCommand = new OleDbCommand($"SELECT * FROM {NameTable}", Connection);
            waitoperation.Fill(celloperation);

            DataTable table = celloperation.Tables[0];
            int low = 0;
            int higth = table.Rows.Count - 1;
            int mid = 0;
            int value = 0;

            while (low <= higth)
            {
                mid = (low + higth) / 2;
                value = (int)table.Rows[mid][0];

                if (value == id)
                {
                    row = table.Rows[mid];
                    return;
                }

                if (value < id)
                    low = mid + 1;
                else
                    higth = mid - 1;
            }
        }
        /// <summary>
        /// Открывает поток редактирования
        /// </summary>
        /// <param name="NameTable">Название таблицы</param>
        /// <param name="val">Парамантр поиска записи</param>
        /// <param name="row">Строка редактирования</param>
        public void BeginUpdateDataRow(string NameTable, object val, ref DataRow row)
        {
            if (Connection.State == ConnectionState.Closed) TryOpen();

            waitoperation.SelectCommand = new OleDbCommand($"SELECT * FROM {NameTable}", Connection);
            waitoperation.Fill(celloperation);

            DataTable table = celloperation.Tables[0];
            foreach (DataRow i in table.Rows)
            {
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    if (!i[j].Equals(val)) continue;
                    row = i;
                    return;
                }
            }
        }
        /// <summary>
        /// Завершает поток редактирования
        /// </summary>
        public void EndUpdateDataRow()
        {
            if (Connection.State != ConnectionState.Open) return;
            if (waitoperation == null || celloperation == null) return;

            try
            {
                OleDbCommandBuilder builder = new OleDbCommandBuilder(waitoperation);
                waitoperation.Update(celloperation);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка обновления: " + ex.Message, "Система", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                TryClose();
                waitoperation.SelectCommand.Dispose();
                celloperation.Dispose();
                celloperation = new DataSet();
            }

        }
        /// <summary>
        /// Обновляет все Id в таблице базы данных
        /// </summary>
        /// <param name="NameTable">Название таблицы</param>
        public void RefreshIdDataTable(string NameTable)
        {
            bool closed = Connection.State == ConnectionState.Closed;
            if (closed) TryOpen();

            OleDbDataAdapter adapter = new OleDbDataAdapter($"SELECT * FROM {NameTable}", Connection);
            DataSet data = new DataSet();
            DataTable table = null;

            try
            {
                adapter.Fill(data);
                table = data.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка операции: " + ex.Message, "Система", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (closed) TryClose();
                return;
            }

            for (int i = 0; i < table.Rows.Count; i++)
            {
                if ((int)table.Rows[i][0] == (i + 1)) continue;
                table.Rows[i][0] = i + 1;
            }

            try
            {
                OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);
                adapter.Update(data);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка удаления: " + ex.Message, "Система", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (closed) TryClose();
            }
        }
    }
}
