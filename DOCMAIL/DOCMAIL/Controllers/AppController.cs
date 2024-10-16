﻿using DOCMAIL.Services.Business;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Web.Mvc;

namespace DOCMAIL.Controllers
{
    public class AppController : Controller
    {
        private readonly ProcessingService processingService = new ProcessingService();
        private readonly InvoiceService invoiceService = new InvoiceService();

        public ActionResult Main()
        {
            return View();
        }

        /// <summary>
        /// Esta funcion se encarga de llamar a listar todos los invoices que cumplen la misma condicion que tipo de registro,y opcionalmente si se pasa el nroRegistro, usa Like para comparar el nro de Invoice. Posterioriomente se le aplica formato Json a los registros.
        /// <param name="tipoRegistro">Tipo de registro a buscar (0 = PENDIENTES // 1 = ENVIADOS )</param>
        /// <param name="nroInvoice">Nro parcial o completo del Commercial Invoice a buscar. Si este valor es "" entonces se listaran todos los invoices segun el valor de tipoRegistro</param>
        /// </summary>
        /// <returns>JsonResult</returns>
        [HttpGet]
        public JsonResult GetInvoices(string tipoRegistro,string nroInvoice)
        {
            var dt = processingService.Listar(tipoRegistro,nroInvoice);

            var RecordsList = dt.AsEnumerable().Select(row => new
            {
                FACTURA = row.Field<string>("FACTURA"),
                FECHA = row.Field<string>("FECHA"),
                DESTINATARIO = row.Field<string>("DESTINATARIO"),
                DIRECCION = row.Field<string>("DIRECCION"),
                MAIL = row.Field<string>("MAIL"),
                USUARIO = row.Field<string>("USUARIO"),
                FECHAULTENV = row.Field<DateTime?>("FECHAULTENV"),
                CANT_ENVIOS = row.Field<string>("CANT_ENVIOS")
            }).ToList();

            // Devuelve la lista como JSON
             return Json(RecordsList, JsonRequestBehavior.AllowGet);

        }

        /// <summary>
        /// Se encarga de llamar al procesamiento y dado de alta de Commercial Invoices traidos desde SAP
        /// </summary>
        /// <returns></returns>
        public void ProcessComercialInvoices()
        {
            processingService.Procesar();
        }

        /// <summary>
        /// Se encarga de llamar al procesamiento y actualiza las invoices seleccionadas segun el tipoRegistro
        /// </summary>
        /// <returns></returns>
        public void UpdateInvoices()
        {
            processingService.Procesar();
        }

        /// <summary>
        /// Se encarga de llamar a la construccion, cambio de cantidad de envios y descarga de un Invoice especifico
        /// </summary>
        /// <returns></returns>
        public ActionResult DescargarInvoice(string nroInvoice)
        {
            var document = invoiceService.RetornarInvoice(nroInvoice);
            processingService.CambiarCategoriaInvoice(nroInvoice);
            using (var stream = new MemoryStream())
            {
                document.Save(stream); 
                stream.Position = 0; 
                return File(stream.ToArray(), "application/pdf", "invoice.pdf");
            }
        }
    }
}