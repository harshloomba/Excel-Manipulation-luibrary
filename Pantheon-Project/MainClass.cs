using System;
using System.Data;
using MySql.Data.MySqlClient;
using log4net;
using log4net.Config;
namespace PantheonProject
{
	class DatabaseConnection
	{
			private static readonly ILog logger =LogManager.GetLogger(typeof(DatabaseConnection));
		public static void Main(string[] args)
		{
			try
			{
				String s = Environment.GetEnvironmentVariable("tab");
				logger.Info("enviornment set value: " + s);


				PantheonProject.PSIManipulation.readCodesDataIntoDataBase(s);
					for (int i = 0; i < 6; i++)
				PantheonProject.PSIManipulation.updateCodeData(i, s);
			
			/*	string connectionString =
			 "Server=localhost;Database=pantheon;Pooling=false;User ID=root;Password=World@1234";
				IDbConnection dbcon;
				dbcon = new MySqlConnection(connectionString);
				dbcon.Open();
				IDbCommand dbcmd = dbcon.CreateCommand();
				// requires a table to be created named employee
				// with columns firstname and lastname
				// such as,
				//        CREATE TABLE employee (
				//           firstname varchar(32),
				//           lastname varchar(32));
				string sql =
					"SELECT AccountNumber,AccountName " +
					"FROM ACCOUNT";
				dbcmd.CommandText = sql;
				IDataReader reader = dbcmd.ExecuteReader();
				while (reader.Read())
				{
					string AccountNumber = (string)reader["AccountNumber"];
					string AccountName = (string)reader["AccountName"];
					Console.WriteLine(AccountNumber + " " + AccountName);
				}
				// clean up
				reader.Close();
				reader = null;
				dbcmd.Dispose();
				dbcmd = null;
				dbcon.Close();
				dbcon = null;*/

				//	GetAccountCountFromMySQL();
			}
			catch (Exception e)
			{
				logger.Info(e.Message);
			}
		}
	}
}
