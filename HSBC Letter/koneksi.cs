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
        //static string constr = Properties.Settings.Default.KonekSql.ToString();

        //Server=192.168.10.13; Initial Catalog=SINARMAS; User ID=Ivan.Witono; Password=Ivan1234!
        //static string constr = Properties.Settings.Default.KonekLocal.ToString();
        //static string constr = "Server=; Initial Catalog=Manulife_PAYDI; User ID=Ivan.Witono; Password=Ivan1234!";
        //static SqlTransaction tran;
        //string vNoman = "";
        //static string constr = @"Data Source=LAPTOP-S496JAR6\DBSIDIQ; Initial Catalog=Manulife_PAYDI; Integrated Security=True"; //User ID=Ivan.Witono;Password=Ivan1234!";
        //static string constr = "Data Source=192.168.10.13; Initial Catalog=AllianzAHCS; User ID=Ivan.Witono; Password=Ivan1234!";

        //static string constr = @"Data Source=LAPTOP-S496JAR6\DBSIDIQ;Initial Catalog = Manulife_DAYCOMMS; Integrated Security = True";
        static string constr = @"Data Source= LAPTOP-S496JAR6\DBSIDIQ;Initial Catalog = HSBC_Letter; Integrated Security = True";
       

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
