using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Forms;

namespace DataBaseWorker
{
    public static class DataBaseAplication
    {
        public static string ConnectionString { get; set; } = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
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
        public static bool TryOpen(string connection)
        {
            ConnectionString = connection;
            return TryOpen();
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
    }
}
