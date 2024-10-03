using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;

namespace COMERCIO_EXTERIOR.Models
{
    public class DMDAL
    {
        /// <summary>
        /// Esta funcion hace ...
        /// </summary>
        /// <returns>Datatable</returns>
        public static DataTable ListRecords()
        {
            DataTable result = new DataTable();

            using (Conexao conn = new Conexao())
            {
                result = conn.Executar_select("SELECT * FROM DMMAUSUA", "COMEX");

            }

            return result;
        }
        /// <summary>
        /// Esta funcion hace ...
        /// </summary>
        /// <returns>Datatable</returns>
        public static DataTable ListarInvoces(string tipo)
        {
            DataTable result = new DataTable();

            using (Conexao conn = new Conexao())
            {
                conn.Criar_parametro("@ENVIADA", tipo);
                //conn.Criar_parametro("@NCIN", "973001");
                result = conn.Executar_proc("DOCMAIL_CONSULTA_PENDIENTES", "COMEX");

            }

            return result;
        }
        /// <summary>
        /// Esta funcion hace ...
        /// </summary>
        /// <returns>Datatable</returns>
        public static List<dynamic> ConseguirUsuarios() {

            List<dynamic> result = new List<dynamic>();

            using (Conexao conn = new Conexao())
            {
                result = conn.DataTableToListDynamic(conn.Executar_select("SELECT USUAUSAP,USUAPASW FROM DMMAUSUA ", "COMEX"));
            }

            return result;
        }

    }
}