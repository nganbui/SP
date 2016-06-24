using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NYDOE.PMO.DAL
{
    public class DataLayer
    {
        #region Private Fields
        /// <summary>
        /// Domain name of the AD Store.
        /// </summary>
        private string _connstring = string.Empty;

        #endregion
        public DataLayer()
        {
            _connstring = "data source=44VSPDEV\\SP2013DEV;integrated security=true;database=Projector Notification"; 
        }
        public DataLayer(string connstring)
        {
            _connstring = connstring;
        }
        
        public string DbConnStr
        {
            get
            {

                return _connstring;
            }

        }

        public DataTable GetDataTableBySP(string sp)
        {

            string err = string.Empty;
            DataTable dt = new DataTable();

            SqlConnection con = new SqlConnection(DbConnStr);
            SqlCommand cmd = new SqlCommand(sp, con);

            cmd.CommandType = CommandType.StoredProcedure;

            try
            {
                con.Open();

                SqlDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                dt.Load(dr);

            }
            catch (Exception ex)
            {
                err = ex.Message;

            }
            return dt;

        }

        public string ExecuteSclrBySP(string sp, SortedList sl)
        {

            string err = string.Empty;
            string returnSclr = "-1";

            SqlConnection con = new SqlConnection(DbConnStr);
            SqlCommand cmd = new SqlCommand(sp, con);

            cmd.CommandType = CommandType.StoredProcedure;

            foreach (DictionaryEntry entry in sl)
            {
                cmd.Parameters.Add(new SqlParameter(entry.Key.ToString(), entry.Value));
            }



            try
            {
                con.Open();
                returnSclr = Convert.ToString(cmd.ExecuteScalar());

            }
            catch (Exception ex)
            {
                err = ex.Message;
            }
            finally
            {
                con.Close();
            }


            return returnSclr;

        }

        public string ExecuteSclrBySPDelete(string sp)
        {

            string err = string.Empty;
            string returnSclr = "-1";

            SqlConnection con = new SqlConnection(DbConnStr);
            SqlCommand cmd = new SqlCommand(sp, con);

            cmd.CommandType = CommandType.StoredProcedure;

            try
            {
                con.Open();
                returnSclr = Convert.ToString(cmd.ExecuteScalar());

            }
            catch (Exception ex)
            {
                err = ex.Message;
            }
            finally
            {
                con.Close();
            }


            return returnSclr;

        }

    }
}
