using DOCMAIL.Models;
using DOCMAIL.Services.Business;
using System.Web.Mvc;

namespace DOCMAIL.Controllers
{
    public class LoginController : Controller
    {

        private readonly SecurityService securityService = new SecurityService();
        public ActionResult Index()
        {
            return View("Login");
        }

        /// <summary>
        /// Esta función se encarga de llamar al securityService para verificar que las credenciales provistas pertenezcan a la BD, si el service devuelve un valor negativo hace saltar el modal indicando que las credenciales no se encuentran registradas.
        /// </summary>
        /// <param name="user">Objeto modelado en base a la clase userModel generado con las credenciales introducidas en el form del Login</param>
        /// <returns>ActionResult</returns>
        public ActionResult Login(UserModel user)
        {
            bool success = securityService.Authenticate(user);

            if (success)
            {
                TempData["LoginFailed"] = false; 
                return RedirectToAction("Main", "App");
            }
            else
            {
                TempData["LoginFailed"] = true; 
                return View("Login", user); 
            }
        }


    }
    
}   