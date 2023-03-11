using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HSBC_Letter
{
    class koneksi
    {
        SqlConnection con;
        public SqlCommand cmd;
        SqlDataAdapter adapter;
        static string constr = Properties.Settings.Default.KonekSql.ToString();

        
       

        // providerName="System.Data.SqlClient";
        DataTable dt;

        public void connection()
        {
            con = new SqlConnection(constr);

            if (con.State != ConnectionState.Closed)
            {
                con.Close();
            }
        }
        public void VerifyDir(string pathdir)
        {
            if (!Directory.Exists(pathdir))
            {
                Directory.CreateDirectory(pathdir);
            }
        }

        public DataTable openTable(string query)
        {
            try
            {
                //select
                con = new SqlConnection(constr);
                if (con.State != ConnectionState.Closed)
                {
                    con.Close();
                }

                con.Open();

                dt = new DataTable();
                adapter = new SqlDataAdapter(query, constr);
                adapter.Fill(dt);

                con.Close();
                return dt;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return new DataTable();
            }
        }

        public void executeQuery(string query)
        {
            con = new SqlConnection(constr);
            if (con.State != ConnectionState.Closed)
            {
                con.Close();
            }

            con.Open();

            cmd = con.CreateCommand();
            cmd.CommandText = query;
            //reader = cmd.ExecuteReader();
            cmd.ExecuteNonQuery();
            con.Close();
        }
    }
}
