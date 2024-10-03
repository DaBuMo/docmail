using System.Web;
using System.Web.Mvc;

namespace COMERCIO_EXTERIOR
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
