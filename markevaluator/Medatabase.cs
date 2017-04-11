using System;
using System.Collections.Generic;
using MySql.Data.MySqlClient;

namespace markevaluator
{
    class Medatabase
    {
        static private MySqlConnection connection;
        static private string connection_string;
        static private MySqlCommand mysql_command;

        /// <summary>
        /// Call this method to open a connection to database;
        /// </summary>
        public static void OpenConnection()
        {
            try
            {
                //Check if connection is already open.
                if (IsConnectionOpen())
                    return;

                //attempt to open connection
                connection_string = String.Format("server={0};uid={1};pwd={2};database={3}", "localhost", "root", "", "medatabase");
                connection = new MySqlConnection(connection_string);
                connection.Open();
            }
            catch (MySqlException mse)
            {
                string errorMessage = "";
                switch(mse.Number)
                {
                    case 1042:
                        errorMessage = "Cannot connect to the specified host may be becuase the MySql service is not running!" + Environment.NewLine + "Please start the service and try again.";
                        break;
                    case 1044:
                        errorMessage = "Access denied for the given username!";
                        break;
                    case 1045:
                        errorMessage = "Access denied for the given username and password!";
                        break;
                    case 1049:
                        errorMessage = "Unknown database!";
                        break;
                    default:
                        errorMessage = "There was an error connecting to database!";
                        break;
                }
                LogWriter.WriteError("While opening MySqlConnection", mse.Message);
                throw new databaseException(errorMessage,mse);
            }
            catch (Exception ex)
            {
                LogWriter.WriteError("While opening MySqlConnection", ex.Message);
                throw new databaseException("Internal error occured!", ex);
            }
        }

        /// <summary>
        /// This method must be called for closing the connection
        /// </summary>
        public static void CloseConnection()
        {
            try
            {
                //check if connection object is null or connection is already closed.
                if (null == connection || connection.State == System.Data.ConnectionState.Closed)
                    return;
                else
                    connection.Close();
            }
            catch (Exception ex)
            {
                LogWriter.WriteError("While closing MySqlConnection", ex.Message);
                throw new databaseException("Internal error occured!", ex);
            }
        }

        /// <summary>
        /// This method will execute a sql statement.
        /// </summary>
        /// <param name="command"></param>
        /// <returns>[int] Returns numbers of rows affected</returns>
        public static int ExecuteQuery(string command)
        {
            int rows_affected = 0;
            try
            {
                //If connection is not open
                if (!IsConnectionOpen())
                {
                    OpenConnection();
                }

                //attempt to execute query
                mysql_command = new MySqlCommand(command, connection);
                rows_affected = mysql_command.ExecuteNonQuery();
            }
            catch (MySqlException mse)
            {
                LogWriter.WriteError("While executing mysqlquery", mse.Message);
                throw new databaseException("An error occured while trying to execute query!", mse);
            }
            catch (Exception ex)
            {
                LogWriter.WriteError("While executing mysqlquery", ex.Message);
                throw new databaseException("Internal error occured!", ex);
            }
            return rows_affected;
        }

        /// <summary>
        /// Fetch records from tables
        /// </summary>
        /// <param name="command">SQL SELECT query</param>
        /// <returns>Row collection</returns>
        public static List<Row> fetchRecords(String command)
        {
            List<Row> rows = new List<Row>();
            try
            {
                Row row;
                //If connection is not open
                if (!IsConnectionOpen())
                    OpenConnection();

                //attempt to execute query
                mysql_command = new MySqlCommand(command, connection);
                MySqlDataReader reader = mysql_command.ExecuteReader();

                while(reader.Read())
                {
                    row = new Row();
                    for (int i = 0; i < reader.FieldCount; i++)
                        row.column.Add(reader.GetName(i),reader[i]);
                    rows.Add(row);
                }

                reader.Close();
            }
            catch (databaseException de)
            {
                LogWriter.WriteError("While fetching records mysql", de.InnerException.Message);
                throw new databaseException("Not connected to database!", de);
            }
            catch (MySqlException mse)
            {
                LogWriter.WriteError("While fetching records mysql", mse.Message);
                throw new databaseException("Something went wrong during a db operation!", mse);
            }
            catch (Exception ex)
            {
                LogWriter.WriteError("While fetching records mysql", ex.Message);
                throw new databaseException("Internal error occured!", ex);
            }
            return rows;
        } 

        /// <summary>
        /// Every method of this class must call this method to check if the connection is open
        /// then perform rest of the operation.
        /// This method can be called to check connection state.
        /// </summary>
        /// <returns>connection state. true{connected}, false{disconnection}</returns>
        public static bool IsConnectionOpen()
        {
            bool connection_state = false;

            if (connection == null)
                connection_state = false;
            else if (connection.State == System.Data.ConnectionState.Open)
                connection_state = true;
            else
                connection_state = false;

            return connection_state;
        }

        /// <summary>
        /// This method will return true if the passed string value is already
        /// present in the database for the given table and column.
        /// </summary>
        /// <param name="table">The name of table</param>
        /// <param name="column">A column must be of string</param>
        /// <param name="value">A string value to check</param>
        /// <returns>True if found</returns>
        public static bool isPresent(string table, string column, string value)
        {
            bool result = false;
            try
            {
                List<Row> rows = fetchRecords("SELECT " + column + " FROM " + table + " WHERE " + column + "='" + value + "'");
                if (rows.Count > 0)
                    result = true;
            }
            catch(Exception ex)
            {
                LogWriter.WriteError("In Medatabase.isPresent() function", ex.Message);
                throw new databaseException("Error ouccred while checking for matching string");
            }
            return result;
        }

        /// <summary>
        /// This method will return true if the passed string value like value is already
        /// present in the database for the given table and column. [SQL LIKE]
        /// </summary>
        /// <param name="table">The name of table</param>
        /// <param name="column">A column must be of string</param>
        /// <param name="value">A string value to compair</param>
        /// <returns>True if found</returns>
        public static bool isPresentLike(string table, string column, string value)
        {
            bool result = false;
            try
            {
                List<Row> rows = fetchRecords("SELECT " + column + " FROM " + table + " WHERE " + column + " LIKE '%" + value + "%'");
                if (rows.Count > 0)
                    result = true;
            }
            catch (Exception ex)
            {
                LogWriter.WriteError("In Medatabase.isPresent(string) function", ex.Message);
                throw new databaseException("Error ouccred while checking for matching string");
            }
            return result;
        }

        /// <summary>
        /// This method will return true if the passed numeric[long] value is already
        /// present in the database for the given table and column. [SQL LIKE]
        /// </summary>
        /// <param name="table">The name of table</param>
        /// <param name="column">A column must be of string</param>
        /// <param name="value">A long or int value to check</param>
        /// <returns>True if found</returns>
        public static bool isPresent(string table, string column, long value)
        {
            bool result = false;
            try
            {
                List<Row> rows = fetchRecords("SELECT " + column + " FROM " + table + " WHERE " + column + "=" + value);
                if (rows.Count > 0)
                    result = true;
            }
            catch(Exception ex)
            {
                LogWriter.WriteError("In Medatabase.isPresent(numeric) function", ex.Message);
                throw new databaseException("Error ouccred while checking for matching number");
            }
            return result;
        }
    }

    class databaseException : Exception
    {
        public databaseException() {
            //to do something extra here
        }

        public databaseException(string message) : base(message) {
            //to do something extra here
        }

        public databaseException(string message, Exception inner) : base(message, inner) {
            //to do something extra here
        }
    }
}
