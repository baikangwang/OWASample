using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using Orange.OWA.HttpWeb;

namespace Orange.OWA.Authentication
{
    public class AuthenticationManager
    {
        private IList<Cookie> _cookieCache = null;
        private static AuthenticationManager _mgr=new AuthenticationManager();

        public IList<Cookie> CookieCache { get { return _cookieCache ?? (_cookieCache = Authenticate(_host, _userName, _password)); }
        }
        public static AuthenticationManager Current { get { return _mgr; } }

        private string _host;
        private string _userName;
        private string _password;
        private string _emailAddress;
        
        protected AuthenticationManager()
        {
            _cookieCache=new List<Cookie>();
            _host = "webmail.taylorcorp.com";
            _userName = "corp\\bkwang";
            _password = "R8ll#qqO2";
            _emailAddress = "bkwang@nltechdev.com";
            _cookieCache = Authenticate(_host, _userName, _password);
        }

        public string Host { get { return _host; } }
        public string UserName { get { return _userName; } }
        public string EmailAddress { get { return _emailAddress; } }

        protected IList<Cookie> Authenticate(string host,string userName,string password)
        {
            string owaUrl = string.Format("https://{0}/exchweb/bin/auth/owaauth.dll", host);
            string desUrl = string.Format("https://{0}/exchange",host);
            string query = string.Format("destination={0}&flags=0&forcedownlevel=0&trusted=0&username={1}&password={2}&SubmitCreds=Log On;", desUrl, userName, password);
            byte[] content = Encoding.UTF8.GetBytes(query);
            
            // Create the web OwaRequest:
            OwaRequest owaRequest = OwaRequest.Post(owaUrl, content);

            OwaResponse response = owaRequest.Send();

            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new Exception(string.Format("Message: Connection failed to url ({0}). Response status code: '{1}'.", desUrl, response.StatusCode));
            }

            response.Close();

            return response.Cookies;
        }

    }
}
