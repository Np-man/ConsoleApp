using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleTestApp
{
    class Program
    {
       
        static void Main(string[] args)
        {
            string filePath;
            Console.WriteLine("Enter the path to read Excel file");
            filePath = Console.ReadLine();
            string con =@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+filePath+";" +
                        @"Extended Properties='Excel 8.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    var dataTabel = new DataTable();
                    dataTabel.TableName= "";
                    dataTabel.Load(dr);
                   foreach(var col in dataTabel.Columns)
                    {                        
                        Console.Write(col);
                        Console.Write("\t");
                    }
                   foreach(DataRow row in dataTabel.Rows)
                    {
                        Console.WriteLine();
                        Console.Write("\t");
                        foreach (var item in row.ItemArray)
                        {                           
                            Console.Write(item);
                            Console.Write("\t");
                        }
                    }
                    Console.ReadLine();



                    
             

                    //Create Excel Connection
                    string ConStr;
                    string HDR;
                    HDR = "YES";
                    ConStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
                        + filePath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
                    OleDbConnection cnn = new OleDbConnection(ConStr);

                    //Get data from Excel Sheet to DataTable
                    OleDbConnection Connection = new OleDbConnection(ConStr);
                    Connection.Open();
                    OleDbCommand oconn = new OleDbCommand("select * from [Sheet2 $]", Connection);
                    OleDbDataAdapter adp = new OleDbDataAdapter(oconn);
                    DataTable dt = new DataTable();
                    adp.Fill(dt);
                    Connection.Close();

                    connection.Open();
                    //Load Data from DataTable to SQL Server Table.
                    using (SqlBulkCopy BC = new SqlBulkCopy(ConStr))
                    {
                       // BC.DestinationTableName = Tracker + "." + tblCustomer;
                        foreach (var column in dt.Columns)
                            BC.ColumnMappings.Add(column.ToString(), column.ToString());
                        BC.WriteToServer(dt);
                    }
                    connection.Close();
                }
            }
        }
        
    }
}
