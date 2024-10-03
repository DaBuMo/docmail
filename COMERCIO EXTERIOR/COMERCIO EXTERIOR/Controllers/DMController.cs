using COMERCIO_EXTERIOR.Models;
using Microsoft.Ajax.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web.Mvc;

namespace COMERCIO_EXTERIOR.Controllers
{
    public class DMController : Controller
    {
        private List<dynamic> users = new List<dynamic>();

        /// <summary>
        /// Esta funcion hace ...
        /// </summary>
        /// <returns>Datatable</returns>
        public ActionResult DataComex()
        {
            users = DMDAL.ConseguirUsuarios();
            return View();
        }
        /// <summary>
        /// Esta funcion hace ...
        /// </summary>
        /// <returns>Datatable</returns>
        [HttpGet]
        public JsonResult GetInvoices(string tipo)
        {
            var dt = DMDAL.ListarInvoces(tipo);


            var RecordsList = dt.AsEnumerable().Select(row => new
            {
                FACTURA = row.Field<string>("FACTURA"),
                FECHA = row.Field<string>("FECHA"),
                DESTINATARIO = row.Field<string>("DESTINATARIO"),
                DIRECCION = row.Field<string>("DIRECCION"),
                MAIL = row.Field<string>("MAIL"),
                USUARIO = row.Field<string>("USUARIO"),
                FECHAULTENV = row.Field<DateTime?>("FECHAULTENV"),
                CANT_ENV = row.Field<string>("CANT_ENVIOS")
            }).ToList();

            // Devuelve la lista como JSON
            return Json(RecordsList, JsonRequestBehavior.AllowGet);
            
        }
        /// <summary>
        /// Esta funcion hace ...
        /// </summary>
        /// <returns>Datatable</returns>
        [HttpGet]
        public JsonResult GetRecords()
        {
            var dt = DMDAL.ListRecords();

            // Convierte el DataTable en una lista de objetos anónimos
            var RecordsList = dt.AsEnumerable().Select(row => new
            {
                USUAUSAP = row.Field<string>("USUAUSAP"),
                USUAPASW = row.Field<string>("USUAPASW"),
                USUANOMB = row.Field<string>("USUANOMB"),
                USUAMAIL = row.Field<string>("USUAMAIL"),
                USUACOIN = row.Field<bool>("USUACOIN"),
                USUASOCO = row.Field<bool>("USUASOCO"),
                USUAADMI = row.Field<bool>("USUAADMI"),
                FECHA = row.Field<string>("FECHA"),
                DESTINATARIO = row.Field<string>("DESTINATARIO"),
                DIRECCION = row.Field<string>("DIRECCION"),
                MAIL = row.Field<string>("MAIL"),
                USUARIO = row.Field<string>("USUARIO"),
                FECHAULTENV = row.Field<DateTime?>("FECHAULTENV"),
                CANT_ENVIOS = row.Field<int>("CANT_ENVIOS")
            }).ToList();


            // Devuelve la lista como JSON
            return Json(RecordsList, JsonRequestBehavior.AllowGet);
        }
    }

   }
