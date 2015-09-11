using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Office365FromAddin.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Office365FromAddin.Utils
{
    public class TokenHelper
    {
        public static string GetAuthorizationRedirect(string resource)
        {
            return String.Format("{0}oauth2/authorize?response_type=code&client_id={1}&resource={2}&redirect_uri={3}OAuth/AuthCode/signalr_id/", 
                SettingsHelper.AzureADAuthority, SettingsHelper.ClientId, resource, SettingsHelper.AppBaseUrl);
        }

        public async static Task<Token> GetAccessTokenWithCode(string signalrRef, string code, string resource)
        {
            //Retrieve access token using authorization code
            Token token = null;
            HttpClient client = new HttpClient();
            string redirect = SettingsHelper.AppBaseUrl + "OAuth/AuthCode/";
            HttpContent content = new StringContent(String.Format(@"grant_type=authorization_code&redirect_uri={0}{1}/&client_id={2}&client_secret={3}&code={4}&resource={5}", 
                redirect, signalrRef, SettingsHelper.ClientId, 
                HttpUtility.UrlEncode(SettingsHelper.ClientSecret), 
                code, resource));
            content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/x-www-form-urlencoded");
            using (HttpResponseMessage response = await client.PostAsync("https://login.microsoftonline.com/common/oauth2/token", content))
            {
                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    token = JsonConvert.DeserializeObject<Token>(json);
                }
            }
            return token;
        }

        public async static Task<Token> GetAccessTokenWithRefreshToken(string refreshToken, string resource)
        {
            //Retrieve access token using refresh token
            Token token = null;
            HttpClient client = new HttpClient();
            HttpContent content = new StringContent(String.Format(@"grant_type=refresh_token&refresh_token={0}&client_id={1}&client_secret={2}&resource={3}", refreshToken, SettingsHelper.ClientId, HttpUtility.UrlEncode(SettingsHelper.ClientSecret), resource));
            content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/x-www-form-urlencoded");
            using (HttpResponseMessage response = await client.PostAsync("https://login.microsoftonline.com/common/oauth2/token", content))
            {
                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    token = JsonConvert.DeserializeObject<Token>(json);
                }
            }
            return token;
        }
    }
}
