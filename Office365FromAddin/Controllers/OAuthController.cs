using Microsoft.AspNet.SignalR;
using Office365FromAddin.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace Office365FromAddin.Controllers
{
    public class OAuthController : Controller
    {
        [Route("OAuth/AuthCode/{id}/")]
        public async Task<ActionResult> AuthCode(string id)
        {
            //Request should have a code from AAD and an id that represents the user in the data store
            if (Request["code"] == null)
                return RedirectToAction("Error", "Home", new { error = "Authorization code not passed from the authentication flow" });

            //get access token using the authorization code
            var token = await TokenHelper.GetAccessTokenWithCode(id, Request["code"], SettingsHelper.OutlookResourceId);
            
            //notify the client through the hub with the access token
            var hubContext = GlobalHost.ConnectionManager.GetHubContext<TokenHub>();
            hubContext.Clients.Client(id).oAuthComplete(token.access_token);

            //return view successfully
            return View();
        }
    }
}