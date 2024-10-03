using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using COMERCIO_EXTERIOR.Models;

namespace COMERCIO_EXTERIOR.Controllers
{
    public class AccessController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Enter(string user,string password)
        {
            try
            {



                return Content("1");
            }
            catch (Exception ex)
            {
                return Content("Ocurrio un Error" + ex.Message);
            }
        }

    }
}