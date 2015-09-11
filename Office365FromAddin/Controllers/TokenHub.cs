using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.AspNet.SignalR;

namespace Office365FromAddin.Controllers
{
    public class TokenHub : Hub
    {
        public void OAuthComplete(string client_id, string access_token)
        {
            Clients.Client(client_id).oAuthComplete(access_token);
        }
    }
}