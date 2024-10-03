using System.Collections.Generic;
using System.Data;

namespace DOCMAIL.Models
{
    public class DMDAL
    {
        public static DataTable ListRecords()
        {
            DataTable result = new DataTable();

            using (Conexao conn = new Conexao())
            {
                result = conn.Executar_select("SELECT * FROM DMMAUSUA", "COMEX");
            }

            return result;
        }

        public static List<dynamic> ConseguirUsuarios()
        {

            List<dynamic> result = new List<dynamic>();

            using (Conexao conn = new Conexao())
            {
                result = conn.DataTableToListDynamic(conn.Executar_select("SELECT USUAUSAP,USUAPASW FROM DMMAUSUA", "COMEX"));
            }

            return result;
        }

    }
}