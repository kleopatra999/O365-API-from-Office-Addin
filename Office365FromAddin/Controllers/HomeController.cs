using Office365FromAddin.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Office365FromAddin.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewData["redirect"] = TokenHelper.GetAuthorizationRedirect(SettingsHelper.OutlookResourceId);
            return View();
        }
    }
}