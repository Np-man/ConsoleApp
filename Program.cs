using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
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
            System.Console.WriteLine("Enter the path to read Excel file");
            filePath = Console.ReadLine();
            //string con =@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+filePath+";" +
            //            @"Extended Properties='Excel 8.0;HDR=Yes;'";
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath+";" +
                    @"Extended Properties='Excel 12.0 HDR=Yes;'";

          
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [DailyUpdate$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    var dataTabel = new DataTable();
                    dataTabel.TableName="DailyTracker";
                    dataTabel.Load(dr);
                   foreach(var col in dataTabel.Columns)
                    {                        
                        Console.Write(col);
                        Console.Write("\t");
                    }
                   foreach(DataRow row in dataTabel.Rows)
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
            }
        }
    }
}
