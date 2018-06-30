using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;

namespace GroupEventDetection
{
    public class Business
    {
        #region Declaration

        OleDbConnection Conn;
        OleDbCommand Cmd;
        OleDbDataReader Reader;
        string Conn_Str = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\DB.mdb";

        #endregion

        public Business()
        {
            //Conn = null;
            Conn = new OleDbConnection(Conn_Str);
            Conn.Open();
        }

        private void closeReader()
        {
            if (Reader != null)
                if (Reader.IsClosed == false)
                    Reader.Close();
        }

        public void executeQuery(string Query)
        {
            try
            {
                

                closeReader();

                Cmd = new OleDbCommand(Query, Conn);
                Cmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public string getValue(string Query)
        {
            string rtnVal = "";
            try
            {
                closeReader();

                Cmd = new OleDbCommand(Query, Conn);
                rtnVal = Cmd.ExecuteScalar().ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return rtnVal;
        }

        public DataTable getTable(string Query)
        {
            DataTable rtnVal = new DataTable();
            try
            {
                closeReader();

                Cmd = new OleDbCommand(Query, Conn);
                Reader = Cmd.ExecuteReader();
                rtnVal.Load(Reader);

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return rtnVal;
        }

        public DataSet getDataSet(string Query)
        {
            DataSet rtnVal = new DataSet();
            try
            {
                closeReader();

                OleDbDataAdapter da = new OleDbDataAdapter(Query, Conn);
                da.Fill(rtnVal);
                da = null;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return rtnVal;
        }

        public bool isHasRows(string Query)
        {
            bool rtnVal = false;
            try
            {
                closeReader();

                Cmd = new OleDbCommand(Query, Conn);
                Reader = Cmd.ExecuteReader();
                rtnVal = Reader.HasRows;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return rtnVal;
        }
    }

}
