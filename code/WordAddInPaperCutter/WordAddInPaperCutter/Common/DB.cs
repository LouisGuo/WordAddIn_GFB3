using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace WordAddInPaperCutter.Common
{
    public class DB
    {
        private SqlConnection con;
        //private string DBpath = "E:/gaofenbaodb/gaofenbao.mdb";
        //private string DBpath = AppDomain.CurrentDomain.BaseDirectory + "resource\\db\\gaofenbao.mdb";
        // <summary>
        /// 打开数据库连接
        /// </summary>
        /// <param name="DBpath">数据库路径(包括数据库名)</param>
        private void Open()
        {
            if (con == null)
                con = new SqlConnection(@"Server=.;database=BaofenbaoTest;Integrated Security=True");
            //con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;;Data Source=" + DBpath );
            //con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;;Data Source=" + DBpath + ";Jet OLEDB:Database Password=gfb2513");
            if (con.State == ConnectionState.Closed)
                con.Open();
        }
        /// <summary>
        /// 创建一个命令对象并返回该对象
        /// </summary>
        /// <param name="sqlStr">数据库语句</param>
        /// <param name="file">数据库所在路径</param>
        /// <returns>OleDbCommand</returns>
        private SqlCommand CreateCommand(string sqlStr)
        {
            Open();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sqlStr;
            cmd.Connection = con;
            return cmd;
        }
        /// <summary>
        /// 执行
        /// </summary>
        /// <param name="sqlStr">SQL语句</param>
        /// <param name="file">数据库所在路径</param>
        /// <returns>返回数值当执行成功时候返回true,失败则返回false</returns>
        public bool ExecuteNonQury(string sqlStr)
        {
            SqlCommand cmd = CreateCommand(sqlStr);
            int result = cmd.ExecuteNonQuery();
            if (result == -1 | result == 0)
            {
                cmd.Dispose();
                Close();
                return false;
            }
            else
            {
                cmd.Dispose();
                Close();
                return true;
            }
        }

        public bool ExecuteNonQury(string sqlStr, string[] argNames, object[] args)
        {
            SqlCommand cmd = CreateCommand(sqlStr);
            for (int i = 0; i < argNames.Length; i++)
            {
                cmd.Parameters.AddWithValue(argNames[i], args[i]);
            }

            int result = cmd.ExecuteNonQuery();
            if (result == -1 | result == 0)
            {
                cmd.Dispose();
                Close();
                return false;
            }
            else
            {
                cmd.Dispose();
                Close();
                return true;
            }
        }


        /// <summary>
        /// 执行数据库查询
        /// </summary>
        /// <param name="sqlStr">查询语句</param>
        /// <param name="tableName">填充数据集表格的名称</param>
        /// <param name="file">数据库所在路径</param>
        /// <returns>查询的数据集</returns>
        public DataSet GetDataSet(string sqlStr)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = CreateCommand(sqlStr);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            dataAdapter.Fill(ds);
            cmd.Dispose();
            Close();
            dataAdapter.Dispose();
            return ds;
        }


        /// <summary>
        /// 执行数据库查询
        /// </summary>
        /// <param name="sqlStr">查询语句</param>
        /// <param name="tableName">填充数据集表格的名称</param>
        /// <param name="file">数据库所在路径</param>
        /// <returns>查询的数据表</returns>
        public DataTable GetDataTable(string sqlStr)
        {
            DataTable dt = new DataTable();
            SqlCommand cmd = CreateCommand(sqlStr);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
            dataAdapter.Fill(dt);
            cmd.Dispose();
            Close();
            dataAdapter.Dispose();
            return dt;
        }
        // <summary>
        /// 生成一个数据读取器OleDbDataReader并返回该OleDbDataReader
        /// </summary>
        /// <param name="sqlStr">数据库查询语句</param>
        /// <returns>返回一个DataReader对象</returns>
        public SqlDataReader GetReader(string sqlStr, string file)
        {
            SqlCommand cmd = CreateCommand(sqlStr);
            SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            //CommadnBehavior.CloseConnection是将于DataReader的数据库链接关联起来
            //当关闭DataReader对象时候也自动关闭链接
            return reader;
        }
        /// <summary>
        /// 关闭数据库
        /// </summary>
        public void Close()
        {
            if (con != null)
                con.Close();
            con = null;
        }
    }
}