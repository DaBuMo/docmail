using DOCMAIL.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web.Mvc;


namespace DOCMAIL.Controllers
{
    public class DMController : Controller { 
    

        // GET: DM
        public ActionResult DataComex()
        {
            return View();
        }

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
                USUAADMI = row.Field<bool>("USUAADMI")
            }).ToList();


            // Devuelve la lista como JSON
            return Json(RecordsList, JsonRequestBehavior.AllowGet);
        }


    }
}