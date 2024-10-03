using DOCMAIL.Models;
using DOCMAIL.Services.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DOCMAIL.Services.Business
{
    public class SecurityService
    {

        private readonly SecurityDAO daoService = new SecurityDAO();

        /// <summary>
        /// Esta función se encarga de comunicarse con la BD para conseguir los registros que cumplan con las credenciales almacenadas en el objeto user, para despues verificar si existe algun usuario registrado con dichas credenciales.
        /// </summary>
        /// <param name="user">Objeto modelado en base a la clase userModel generado con las credenciales introducidas en el form del Login</param>
        /// <returns>Bool</returns>
        public bool Authenticate(UserModel user)
        {
            var usuario = daoService.FindByUser(user);
            return usuario.Count == 1;
        }
    }
}