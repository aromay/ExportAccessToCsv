using System;
using System.Data.OleDb;
using System.IO;
using System.Text;

namespace ExportAccessToCSV
{
    static class Program
    {
        static void Main()
        {
            ExportAccessTableToCSV("database.mdb", "tableName", ";", "filename.csv");
        }

        public static void ExportAccessTableToCSV(string accessLocation, string tableName, string separator, string commaSeparatedFilename)
        {
            string connectionString = String.Concat("Provider=Microsoft.Jet.OleDb.4.0;Data Source=", accessLocation, ";");
            string queryString = String.Concat("SELECT * FROM ", tableName);

            try
            {
                OleDbConnection connection = new OleDbConnection(connectionString);
                OleDbCommand command = new OleDbCommand(queryString, connection);

                connection.Open();

                OleDbDataReader dr = command.ExecuteReader();

                int fields = dr.FieldCount - 1;

                StreamWriter sw = new StreamWriter(commaSeparatedFilename);

                while (dr.Read())
                {
                    StringBuilder sb = new StringBuilder();

                    for (int i = 0; i <= fields; i++)
                    {
                        sb.Append(dr[i].ToString() + separator);

                    }

                    sw.WriteLine(sb.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}
