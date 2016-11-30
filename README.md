# Calling Office 365 APIs from an Office add-in

>**Note:**  We will be removing this sample from the site on December 15, 2016. If you’d like to keep a copy of this sample for your own reference, please download or clone the repo.

This repository contains a sample that shows how to connect to the Office 365 APIs from an Office Add-in. This can be accomplished a number of ways, but has the following challenges:

- Office add-ins currently have no concept of identity (even though users sign-in to the modern Office).
- Office add-ins can only display pages from domains that are explicitly registered in the add-in manifest. When an add-in tries to display an unregistered domain, the page will get kicked out to a popup.
- The Office 365 login can be federated by Office 365 customers (usually with Active Directory Federation Services (ADFS)), which makes it impossible to predict all the domains the add-in might need to display.
- Internet Explorer Security Zones can make popup/add-in communication hard to predict.

These challenges were heavy drivers in determining the architecture of this sample. The sample forces the OAuth flow into a popup (rather than trying to avoid the popup). To avoid browser security zones, the add-in leverages web sockets (specifically SignalR) to pass tokens "through the server" back to the add-in. For more information on this pattern, see Richard diZerega's blog post on [Connecting to Office 365 from an Office add-in](http://blogs.msdn.com/b/richard_dizeregas_blog/archive/2015/08/10/connecting-to-office-365-from-an-office-add-in.aspx).

## Setup ##
To setup the application, register a new Azure AD application through the Azure Management Portal or using the Add Connected Service Wizard. The add-in needs "Read Contacts" permissions against Exchange Online.

If you manually register the application in Azure AD, you should update **ida:ClientId** and **ida:ClientSecret** appSettings in the web.config with values from your registration:

    <add key="ida:ClientId" value="a8d33bab-8a04-48a5-83dd-8492e63db131" />
    <add key="ida:ClientSecret" value="iq+5dGcYBLhZhvrzyZ25/dAyIYG3SUmmKiAv2MjHd40=" />

The add-in authenticated against Office 365 and then queries Contacts in Exchange Online, so it is suggested that you add a few contacts to test with.

## Utils/TokenHelper.cs ##
This file contains utility functions for interacting with Azure AD, including GetAuthorizationRedirect, GetAccessTokenWithCode, and GetAccessTokenWithRefreshToken.

## Controllers/OAuthController.cs ##
This is the MVC controller that handles the redirect from Azure AD when requesting an authorization code. It contains a single AuthCode action that uses the code to get an access token and pass it back to the OAuth popup via the SignalR web socket.


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

## Views/Home/Index.cshtml ##
This is the primary view that is displayed in the add-in. It contains all the client-side logic for launching the OAuth popup, listening for access tokens via SignalR, and calling the Office 365 APIs client-side.
