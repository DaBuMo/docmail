using DOCMAIL.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace DOCMAIL.Services.Data
{
    public class SecurityDAO
    {
        /// <summary>
        /// Esta función se encarga de buscar y devolver los registros que cumplan con las credenciales almacenadas en el objeto user
        /// </summary>
        /// <param name="user">Objeto modelado en base a la clase userModel generado con las credenciales introducidas en el form del Login</param>
        /// <returns> Dynamic List con los registros que cumplan </returns>
        internal List<dynamic> FindByUser(UserModel user)
        {
            List<dynamic> result = new List<dynamic>();
            
            using (Conexao conn = new Conexao())
            {
                conn.Criar_parametro_objeto("@Username", user.Username);
                conn.Criar_parametro_objeto("@Password", user.Password);

                result = conn.DataTableToListDynamic(conn.Executar_select("SELECT USUAUSAP,USUAPASW FROM DMMAUSUA WHERE USUAUSAP = @Username AND USUAPASW = @Password", "COMEX"));
            }

            return result;
        }
    }

}