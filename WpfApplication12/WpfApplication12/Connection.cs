using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.DataAccess.Client;
using System.Windows;
using System.Data;  
using System.Data.SqlClient;
using System.Windows.Controls;






namespace WpfApplication12
{
    class Connection
    {
        //define variables
        public string user;
        public string pass;
        public string host;
        public string port;
        public string service;
        OracleConnection con;
        public void OpenDb()
        {
            //create a new connection
            con = new OracleConnection();
            con.ConnectionString = "user id=" + user + ";password=" + pass + ";data source=" +
                                    "(DESCRIPTION=(ADDRESS=(PROTOCOL=tcp)" +
                                    "(HOST=" + host + ")(PORT=" + port + "))(CONNECT_DATA=" +
                                     "(SERVICE_NAME=" + service + ")))";
            //open connection
            try
            {
                con.Open();
                MessageBox.Show("Conectare cu succes");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }

        }

        public void CloseDb()
        {
            // Close and Dispose OracleConnection object
            con.Close();
            con.Dispose();
            MessageBox.Show("Connection closed");
        }

        //public void Sql(string s)
        public DataSet FillDataGrid(string s)
        {
            // Create the OracleCommand
            OracleCommand cmd = new OracleCommand(s);
            cmd.Connection = con;
            cmd.CommandText = s;

            // Execute command, create OracleDataReader object
            // OracleDataReader reader = cmd.ExecuteReader();

            DataSet ds = new DataSet();
            try
            {
                OracleDataAdapter a = new OracleDataAdapter(cmd.CommandText, con);
                a.Fill(ds);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }

            return ds;

            //  Dgv.DataSource(ds.Tables[0]);

            // if (reader.HasRows)
            //{
            //get headerele si scrie un rand
            //  var columns = Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToList();
            //var columns = Enumerable.Range(0, 4).Select(reader.GetName).ToList();

            /* foreach (string coloana in columns)
                 MessageBox.Show(coloana + " ");
                

             while (reader.Read())
             {
                 object[] values = new object[reader.FieldCount];
                 //object[] values = new object[4];

                 int numColumns = reader.GetValues(values); //after "reading" a row

                 for (int i = 0; i < numColumns; i++)
                 {
                     MessageBox.Show(values[i] + " ");
                 }
                    
             }
             */




            // }
        }
        // Dispose OracleDataReader object
        //  reader.Dispose();

        // Dispose OracleCommand object
        // cmd.Dispose();
        
        public  string Sql(string s)
        {
            // Create the OracleCommand
            OracleCommand cmd = new OracleCommand(s);
            cmd.Connection = con;
            cmd.CommandText = s;

            // Execute command, create OracleDataReader object
            try
            {
                OracleDataReader reader = cmd.ExecuteReader();
            
            if (reader.HasRows)
            {
                StringBuilder builder = new StringBuilder();

                //get headerele si scrie un rand
                var columns = Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToList();
                //var columns = Enumerable.Range(0, 4).Select(reader.GetName).ToList();

                foreach (string coloana in columns)
                {
                    builder.Append(coloana);
                    builder.Append(' ');
                }
               builder.Append("\n");
                
                while (reader.Read())
                {
                    object[] values = new object[reader.FieldCount];
                    //object[] values = new object[4];

                    int numColumns = reader.GetValues(values); //after "reading" a row

                    for (int i = 0; i < numColumns; i++)
                    {
                        builder.Append(values[i]);
                        builder.Append(' ');

                    }
                    builder.Append("\n");

                }


                return builder.ToString();  

            }
            else
            {
                int numberOfRecords = cmd.ExecuteNonQuery();
                return numberOfRecords.ToString() + " rows updated";

            }
                 reader.Dispose();

            // Dispose OracleCommand object
            cmd.Dispose();
            }
            catch (Exception e)
                {
                    return e.ToString();
                }
           
           }

           // return "succes";
                // Dispose OracleDataReader object
            
        }

    }

