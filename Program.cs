using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ConsoleTestApp.Globals;



namespace ConsoleTestApp
{
    class Program
    {
        public DataTable getExcelSheet(string path)
        {
            DataTable excelData = new DataTable();
            string excelConString = null, conString;
            excelConString = ConfigurationManager.ConnectionStrings["ExcelConStringxls"].ConnectionString;
            conString = string.Format(excelConString, path);
            try
            {
                using (OleDbConnection connection = new OleDbConnection(conString)) 
                {
                    connection.Open();
                    using (OleDbCommand excelCommand = new OleDbCommand("select * from [DailyUpdate$]", connection))
                    {
                        using (OleDbDataReader data = excelCommand.ExecuteReader())
                        {
                            excelData.TableName = "DailyTracker";
                            excelData.Load(data);
                        }

                    }

                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.ReadLine();
            }

            return excelData;
        }
        /// <summary>
        /// This method displays excel sheet on console
        /// </summary>
        /// <param name="dataTable"></param>
        public void displayExcelOnConsole(DataTable dataTable)
        {
            foreach (var col in dataTable.Columns)
            {
                Console.Write(col);
                Console.Write("\t");
            }
            foreach (DataRow row in dataTable.Rows)
            {
                Console.WriteLine();
                foreach (var item in row.ItemArray)
                {
                    Console.Write(item);
                    Console.Write("\t");
                }
            }
            Console.ReadLine();
        }
        /// <summary>
        /// This method inserts excel sheet values into TaskStatus table
        /// </summary>
        /// <param name="dataTable"></param>
        /// <returns>Returns number or rows affected</returns>
        public int insertExcelToDatabase(DataTable dataTable)
        {
            int rows = 0;
            string conString = ConfigurationManager.ConnectionStrings["SqlConnection"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(conString))
            {
                try
                {
                    string sqlCommand = "insert into taskstatus([taskid],[taskstatus]) values (@taskid,@taskstatus)";
                    connection.Open();
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        using (SqlCommand cmd = new SqlCommand(sqlCommand, connection))
                        {
                            cmd.Parameters.Add("@taskid", SqlDbType.Int).Value = dataTable.Rows[i]["TaskId"];
                            cmd.Parameters.Add("@taskstatus", SqlDbType.NVarChar).Value = dataTable.Rows[i]["TaskStatus"];
                            rows = cmd.ExecuteNonQuery();
                        }

                    }

                    connection.Close();
                    return rows;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    Console.ReadLine();
                    return rows;

                }
            }
        }
        /// <summary>
        /// This method is used to execute user query
        /// </summary>
        /// <param name="query"></param>
        public void executeUserQuery(string query)
        {
            string conString = ConfigurationManager.ConnectionStrings["SqlConnection"].ConnectionString;
            try
            {
                using (SqlConnection conneciton = new SqlConnection(conString))
                {
                    conneciton.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conneciton))
                    {
                        using (SqlDataReader data = cmd.ExecuteReader())
                        {
                            while (data.Read())
                            {
                                for (int i = 0; i < data.FieldCount; i++)
                                {
                                    Console.WriteLine(data.GetValue(i));
                                }
                                Console.WriteLine();
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.ReadLine();
            }
        }

        static void Main(string[] args)
        {
            string filePath;
            string userQuery;
            System.Console.WriteLine("Enter the path to read Excel file");
            filePath = Console.ReadLine();

            Program obj = new Program();//Object for class program

            DataTable excelSheet = obj.getExcelSheet(filePath);//retrieves excelsheet

            obj.displayExcelOnConsole(excelSheet);//displays excelsheet on console

            int rowsAffectd = obj.insertExcelToDatabase(excelSheet);

            System.Console.WriteLine("Enter query to execute");
            //query to read from database is remaining
            userQuery = Console.ReadLine();

            obj.executeUserQuery(userQuery);

        }
    }
}
