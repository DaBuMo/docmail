using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Dynamic;

namespace COMERCIO_EXTERIOR.Models
{
    public class Conexao : IDisposable
    {
        private SqlConnection SqlConn;
        public SqlTransaction SqlTrans;
        public SqlCommand SqlComm = new SqlCommand();
        private DataTable dtRetorno = new DataTable();

        public String strRetorno;
        private Boolean blnRetorno;
        private DataTable dtRetornoTrans = new DataTable();

        #region [ Properties ]

        public String TextoRetorno
        {
            //Retorna o texto da procedure para saber qual foi a mensagem de erro ou de sucesso
            get
            {
                return strRetorno;
            }
        }

        public Boolean ResultadoTrans
        {
            //Retorna se a transação foi executada ou não com sucesso
            get
            {
                return blnRetorno;
            }
        }

        public DataTable DataTableTrans
        {
            //Retorna o recordset obtido da transação pois para algumas telas é necessário pegar o código de retorno
            get
            {
                return dtRetornoTrans;
            }
        }
        #endregion

        public void AbreConexao(String pStrConexao = "")
        {
            SqlConn = new SqlConnection();
            SqlConn.ConnectionString = ConfigurationManager.ConnectionStrings[pStrConexao].ToString();
            SqlConn.Open();
            BeginTrans();
        }

        public void Criar_parametro(String pStrName, String pStrValor)
        {
            if (pStrValor != null)
            {
                //Retirar as aspas (') para evitar erros no banco de dados
                pStrValor = pStrValor.Replace("'", "´");
            }

            SqlComm.Parameters.AddWithValue(pStrName, pStrValor).Direction = ParameterDirection.Input;
        }

        public void Criar_parametro_objeto(String pStrName, object pStrValor)
        {
            SqlComm.Parameters.AddWithValue(pStrName, pStrValor).Direction = ParameterDirection.Input;
        }

        public void Executar_comando(String pStrComando, String pStrConexao = "")
        {
            //Nesta função pode-se passar um comando que não precise de retorno
            try
            {
                SqlConnection SqlConn = new SqlConnection();
                SqlConn.ConnectionString = ConfigurationManager.ConnectionStrings[pStrConexao].ToString();
                SqlConn.Open();
                SqlComm.Connection = SqlConn;

                SqlComm.CommandType = CommandType.Text;
                SqlComm.CommandText = pStrComando;

                SqlComm.ExecuteNonQuery();

                SqlConn.Close();

                blnRetorno = true;
            }
            catch (Exception ex)
            {
                blnRetorno = false;
                strRetorno = ex.Message.ToString();
            }
        }

        public DataTable Executar_select(String pStrComando, string pStrConexao = "")
        {
            //Nesta função pode-se passar um código de texto para Seleção ou mesmo um comando qq
            DataTable dtRetorno = new DataTable();
            SqlDataAdapter MyAdapter = new SqlDataAdapter();

            try
            {
                SqlConnection SqlConn = new SqlConnection();
                SqlConn.ConnectionString = ConfigurationManager.ConnectionStrings[pStrConexao].ToString();
                SqlComm.Connection = SqlConn;

                SqlComm.CommandType = CommandType.Text;
                SqlComm.CommandText = pStrComando;
                MyAdapter.SelectCommand = SqlComm;
                MyAdapter.Fill(dtRetorno);

                blnRetorno = true;
            }
            catch (Exception ex)
            {
                blnRetorno = false;
                strRetorno = ex.Message.ToString();
            }

            return dtRetorno;
        }

        public DataTable Executar_proc(String pStrComando, string pStrConexao = "")
        {
            //Aqui se executa uma procedure para obter o retorno, NÃO aceita uma string de Select
            DataTable dtRetorno = new DataTable();
            SqlDataAdapter MyAdapter = new SqlDataAdapter();

            try
            {
                SqlConnection SqlConnection1 = new SqlConnection();
                SqlConnection1.ConnectionString = ConfigurationManager.ConnectionStrings[pStrConexao].ToString();
                SqlComm.Connection = SqlConnection1;

                SqlComm.CommandType = CommandType.StoredProcedure;
                SqlComm.CommandText = pStrComando;
                SqlComm.CommandTimeout = 0;
                MyAdapter.SelectCommand = SqlComm;
                MyAdapter.Fill(dtRetorno);

                blnRetorno = true;
            }
            catch (Exception ex)
            {
                blnRetorno = false;
                strRetorno = ex.Message.ToString();
            }

            SqlComm.Parameters.Clear(); //Limpar os parâmetros

            return dtRetorno;
        }

        public DataTable Executar_proc_trans(String pStrComando)
        {
            //Executa uma procedure transacional. Só será feito commit se tudo for OK
            SqlDataAdapter MyAdapter = new SqlDataAdapter();

            try
            {
                SqlComm.Connection = SqlConn;
                SqlComm.CommandType = CommandType.StoredProcedure;
                SqlComm.CommandTimeout = 120;
                SqlComm.CommandText = pStrComando;
                MyAdapter.SelectCommand = SqlComm;
                //dtRetornoTrans.Clear(); //Limpar a tabela, pois pode haver sujeira da última execução
                MyAdapter.Fill(dtRetornoTrans);

                //Verificar se o retorno do banco de dados foi OK
                if (dtRetornoTrans.Rows[0][0].ToString() == "1")
                {
                    if (dtRetornoTrans.Columns.Count > 1)
                        strRetorno = dtRetornoTrans.Rows[0][1].ToString();
                    else
                        strRetorno = "";

                    blnRetorno = true;
                }
                else
                {
                    strRetorno = dtRetornoTrans.Rows[0][1].ToString();
                    if (strRetorno.ToString() == "OK")
                    {
                        blnRetorno = true;
                    }
                    else
                    {
                        blnRetorno = false;
                    }
                }
                //Tratar os parâmetros para pegar os retornos em que o parâmetro é de OutPut
                SqlComm.Parameters.Clear(); //Limpar os parâmetros
            }
            catch (Exception ex)
            {
                blnRetorno = false;
                strRetorno = ex.Message.ToString();
            }

            return dtRetornoTrans;
        }

        public void Executar_comando_trans(String pStrComando)
        {
            try
            {
                SqlComm.Connection = SqlConn;
                SqlComm.CommandType = CommandType.Text;
                SqlComm.CommandText = pStrComando;
                SqlComm.ExecuteNonQuery();

                blnRetorno = true;
            }
            catch (Exception ex)
            {
                blnRetorno = false;
                strRetorno = ex.Message.ToString();
            }
        }

        public void FecharConexao()
        {
            if (SqlConn.State == ConnectionState.Open)
            {
                SqlConn.Close();
                SqlConn = null;
            }
        }

        public void BeginTrans()
        {
            SqlTrans = SqlConn.BeginTransaction();
            SqlComm.Transaction = SqlTrans;
            blnRetorno = true;
        }

        public void CommitTrans()
        {
            SqlTrans.Commit();
            FecharConexao();
        }

        public void RollBackTrans()
        {
            if (SqlConn != null && SqlConn.State == ConnectionState.Open)
            {
                SqlTrans.Rollback();
                FecharConexao();
            }
        }

        public void Dispose()
        {
            ((IDisposable)dtRetorno).Dispose();
        }

        public List<dynamic> Executar_proc_list(String pStrComando, string pStrConexao = "")
        {
            List<dynamic> retorno = new List<dynamic>();
            DataTable dtRetorno = new DataTable();
            SqlDataAdapter MyAdapter = new SqlDataAdapter();

            try
            {
                SqlConnection SqlConnection1 = new SqlConnection();
                SqlConnection1.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings[pStrConexao].ToString();
                SqlComm.Connection = SqlConnection1;

                SqlComm.CommandType = CommandType.StoredProcedure;
                SqlComm.CommandText = pStrComando;
                SqlComm.CommandTimeout = 0;
                MyAdapter.SelectCommand = SqlComm;
                MyAdapter.Fill(dtRetorno);

                retorno = DataTableToListDynamic(dtRetorno);
                blnRetorno = true;
            }
            catch (Exception ex)
            {
                blnRetorno = false;
                strRetorno = ex.Message.ToString();
                throw ex;
            }

            SqlComm.Parameters.Clear();
            return retorno;
        }

        public List<dynamic> DataTableToListDynamic(DataTable dt)
        {
            var dynamicDt = new List<dynamic>();
            foreach (DataRow row in dt.Rows)
            {
                dynamic dyn = new ExpandoObject();
                dynamicDt.Add(dyn);
                foreach (DataColumn column in dt.Columns)
                {
                    var dic = (IDictionary<string, object>)dyn;
                    dic[column.ColumnName] = row[column];
                }
            }
            return dynamicDt;
        }

        public List<dynamic> Executar_select_list(String pStrComando, string pStrConexao = "")
        {
            //Nesta função pode-se passar um código de texto para Seleção ou mesmo um comando qq
            List<dynamic> retorno = new List<dynamic>();
            DataTable dtRetorno = new DataTable();
            SqlDataAdapter MyAdapter = new SqlDataAdapter();

            try
            {
                SqlConnection SqlConn = new SqlConnection();
                SqlConn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings[pStrConexao].ToString();
                SqlComm.Connection = SqlConn;

                SqlComm.CommandType = CommandType.Text;
                SqlComm.CommandText = pStrComando;
                MyAdapter.SelectCommand = SqlComm;
                MyAdapter.Fill(dtRetorno);

                retorno = DataTableToListDynamic(dtRetorno);
                blnRetorno = true;
            }
            catch (Exception ex)
            {
                blnRetorno = false;
                strRetorno = ex.Message.ToString();
                throw ex;
            }

            return retorno;
        }

        public List<dynamic> DataTableToList(DataTable dataTable)
        {
            try
            {
                return DataTableToListDynamic(dataTable);
            }
            catch
            {
                return null;
            }
        }
    }
}