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

namespace DataBaseWorker
{
    public static class DataBaseAplication
    {
        private const string BaseConnection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
        public static string ConnectionString { get; set; }
        private static OleDbConnection conection = new OleDbConnection();

        public static OleDbConnection Connection { get => conection; }

        public static bool TryOpen()
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
        public static bool TryOpen(string NameBase)
        {
            ConnectionString = BaseConnection + NameBase;
            return TryOpen();
        }

        public static void BuildConnection()
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

        public static bool TryClose()
        {
            if(Connection.State == System.Data.ConnectionState.Open)
            {
                Connection.Close();
                return true;
            }

            return false;
        }


        public static DataTable GetDataTable(string NameTable)
        {
            bool closed = Connection.State == ConnectionState.Closed;
            if (closed) TryOpen();

            OleDbDataAdapter adapter = new OleDbDataAdapter($"SELECT * FROM {NameTable}", Connection);
            DataSet data = new DataSet();

            adapter.Fill(data);

            if (closed) TryClose();

            return data.Tables[0];
        }

        public static void AddDataRow(string NameTable, params object[] colums)
        {
            bool closed = Connection.State == ConnectionState.Closed;
            if (closed) TryOpen();

            OleDbDataAdapter adapter = new OleDbDataAdapter($"SELECT * FROM {NameTable}", Connection);
            DataSet data = new DataSet();
            DataTable table = data.Tables[0];

            DataRow row = table.NewRow();

            try
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    if (i < colums.Length)
                        row[i] = colums[i];
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

        public static void RemoveDataRow(string NameTable, object parametr)
        {
            bool closed = Connection.State == ConnectionState.Closed;
            if (closed) TryOpen();

            OleDbDataAdapter adapter = new OleDbDataAdapter($"SELECT * FROM {NameTable}", Connection);
            DataSet data = new DataSet();
            DataTable table = data.Tables[0];
            bool equal = false;

            for(int i = 0; i < table.Rows.Count; i++)
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

        public static void RemoveAtDataRow(string NameTable, int id)
        {
            bool closed = Connection.State == ConnectionState.Closed;
            if (closed) TryOpen();

            OleDbDataAdapter adapter = new OleDbDataAdapter($"SELECT * FROM {NameTable}", Connection);
            DataSet data = new DataSet();
            DataTable table = data.Tables[0];

            for (int i = 0; i < table.Rows.Count; i++)
            {
                if ((int)table.Rows[i][0] != id) continue;
                table.Rows[i].Delete();
                break;
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
