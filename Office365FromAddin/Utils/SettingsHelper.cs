using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace Office365FromAddin.Utils
{
    public class SettingsHelper 
    {
        public static string ClientId {
            get { return ConfigurationManager.AppSettings["ida:ClientId"]; }
        }
        
        public static string ClientSecret {
            get { return ConfigurationManager.AppSettings["ida:ClientSecret"]; }
        }

        public static string AppBaseUrl {
            get { return ConfigurationManager.AppSettings["ida:AppBaseUrl"]; }
        }

        public static string OutlookResourceId {
            get { return "https://outlook.office365.com/"; }
        }

        public static string AzureADAuthority {
            get { return "https://login.microsoftonline.com/common/"; }
        }

        public static string ClaimTypeObjectIdentifier {
            get { return "http://schemas.microsoft.com/identity/claims/objectidentifier"; }
        }
    }
}