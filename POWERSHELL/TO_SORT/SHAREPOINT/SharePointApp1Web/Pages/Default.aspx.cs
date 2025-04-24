using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace SharePointApp1Web.Pages
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            var key = Session["CashKey"];

            // sharepoint url (app hosted url)
            var hostWeb = Page.Request["SPHostUrl"];

            // Sharepoint url with app Title (app deployed sp url)
            Uri SharePointUri = new Uri(hostWeb + "/SharePointApp1/");

            // This is first time the app is running
            if (key == null)
            {
                // get the TokenString
                var contextTokenString = TokenHelper.GetContextTokenFromRequest(Page.Request);
                // Get the contexttoken by passing the token string
                var contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, Request.Url.Authority);

                //Since some browsers does not support cookie name more than 40 chars
                // Im taking first 40 chars
                var cookieName = contextToken.CacheKey.Substring(0, 40);

                //Add User specific cookie name to the Session
                Session.Add("CashKey", cookieName);

                // Get the Refresh Token
                var refreshToken = contextToken.RefreshToken;
                // Add the cookie value (refresh Token)
                Response.Cookies.Add(new HttpCookie(cookieName, refreshToken));

            }
            else
            {
                // USER already in the applicaiton,  means it is not getting redirect from the appRedirect
                // So contextstring is null 
            }
        }


        //Response.Write(val);

        protected void Button1_Click(object sender, EventArgs e)
        {

            // sharepoint url (app hosted url)
            var hostWeb = Page.Request["SPHostUrl"];

            // Sharepoint url with app name (app deployed sp url)
            Uri SharePointUri = new Uri(hostWeb + "/SharePointApp1/");

            // Get the cookie name from the session
            var key = Session["CashKey"] as string;
            // Get the refresh token from the cookie
            var refreshToken = Request.Cookies[key].Value;

            //Get the access Token By pasing refreshToken
            // 00000003-0000-0ff1-ce00-000000000000 is principla name for SP2013 it is unique
            var accessToken = TokenHelper.GetAccessToken(refreshToken,
                "00000003-0000-0ff1-ce00-000000000000",
                SharePointUri.Authority, TokenHelper.GetRealmFromTargetUrl(SharePointUri));

            // Access the Sharepoint Do your work
            using (var clientContext = TokenHelper.GetClientContextWithAccessToken("https://rajee.sharepoint.com/SharepointApp1", accessToken.AccessToken))
            {
                clientContext.Load(clientContext.Web, web => web.Title);
                clientContext.ExecuteQuery();
                Response.Write(clientContext.Web.Title);
            }
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            var hostWeb = Page.Request["SPHostUrl"];
            var val = TokenHelper.GetAppContextTokenRequestUrl(hostWeb, Server.UrlEncode(Request.Url.ToString()));
        }
    }
}