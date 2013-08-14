using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Cache;
using System.Net.Security;
using System.Security;
using System.Security.Authentication;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace OWASample
{
    public class WebDAVRequest
    {
        private  string _hostName = "https://webmail.taylorcorp.com";//"https://mail.mycompany.com";
        private string _url = "https://webmail.taylorcorp.com/owa";
        private string _postUrl = "https://webmail.taylorcorp.com/owa/auth.owa";
        private  string path = "/public/Calendars/Board%20Room";
        private  string _username = "corp\\bkwang";//"user";
        private  string _password = "R8ll#qqO2";//"password";
        //private CookieContainer _cookieContainerPost = null;
        //private CookieContainer _cookieContainerGet = null;
        private int _requestTimeout = 15000;
        private string _successLoginText = "New Message";
        private TimeSpan _responseTime=new TimeSpan();


        /// 
        /// Default WebDAVRequest constructor. Set the certificate callback here.
        /// 

        public WebDAVRequest()
        {
            ServicePointManager.ServerCertificateValidationCallback += CheckCert;
        }

        private string SecureStringToString(SecureString input)
        {
            IntPtr bstr = System.Runtime.InteropServices.Marshal.SecureStringToBSTR(input);
            try
            {
                return System.Runtime.InteropServices.Marshal.PtrToStringBSTR(bstr);
            }
            finally 
            {
                System.Runtime.InteropServices.Marshal.FreeBSTR(bstr);
            }
        }

        /// 
        /// Authenticate against the Exchange server and store the authorization cookie so we can use
        /// it for future WebDAV requests.
        /// 

        public void Authenticate()
        {
            //string authURI = server + "/owa/auth.owa";//"/exchweb/bin/auth/owaauth.dll";

            // Create the web request body:

            //SecureString temp=new SecureString();
            //foreach (char t in password)
            //{
            //    temp.AppendChar(t);
            //}
            //password = SecureStringToString(temp);
            CookieContainer get = OWAGET(this._url);
            CookieContainer post = OWAPOST(this._hostName, this._url, this._postUrl,this._username,this._password, get);

        }

        private CookieContainer OWAGET(string url)
        {
            HttpWebRequest request = (HttpWebRequest)System.Net.WebRequest.Create(url);
            //request.Method = "POST";
            request.Method = "GET";
            request.ContentType = "application/x-www-form-urlencoded";
            request.Timeout = _requestTimeout;
            request.ServicePoint.Expect100Continue = false;
            request.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
            request.CookieContainer = new CookieContainer();
            request.KeepAlive = false;
            request.AllowAutoRedirect = true;

            HttpWebResponse response;
            try
            {
                response = request.GetResponse() as HttpWebResponse;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Message: Coonnection failed to url ({0}). {1}",url,ex.Message));
            }

            if (response.StatusCode != HttpStatusCode.OK)
            {

                throw new Exception(string.Format("Message: Connection failed to url ({0}). Response status code: '{1}'.", url, response.StatusCode));
            }

            using (Stream stream =response.GetResponseStream())
            {
                using (StreamReader sr=new StreamReader(stream) )
                {
                    string result = sr.ReadToEnd();
                    sr.Close();
                }
                stream.Close();
            }

            CookieContainer getCookieContainer = new CookieContainer();
            foreach (Cookie cookie in response.Cookies)
            {
                getCookieContainer.Add(cookie);
            }

            response.Close();

            return getCookieContainer;
        }

        private CookieContainer OWAPOST(string hostName, string url,string postUrl,string userName,string password, CookieContainer getCookieContainer)
        {
            // Create the web request:

            HttpWebRequest request = (HttpWebRequest)System.Net.WebRequest.Create(postUrl);
            request.Method = "POST";
            //request.Method = "GET";
            request.ContentType = "application/x-www-form-urlencoded";
            request.Timeout = _requestTimeout;
            request.ServicePoint.Expect100Continue = false;
            request.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
            //request.CookieContainer = new CookieContainer();
            request.KeepAlive = false;
            request.AllowAutoRedirect = true;

            Uri uri = new Uri(_url);
            CookieContainer postCookieContainer = new CookieContainer();
            Cookie outlookJavascriptCookie = null;
            if (getCookieContainer.Count > 0)
            {
                Cookie outlookSessionCookie = getCookieContainer.GetCookies(uri)[0];
                outlookJavascriptCookie=new Cookie("PBack","0",outlookSessionCookie.Path,outlookSessionCookie.Domain);
            }
            else
            {
                outlookJavascriptCookie = new Cookie("PBack", "0", "/", hostName);
            }
            postCookieContainer.Add(outlookJavascriptCookie);
            request.CookieContainer = postCookieContainer;

            //string postContent = string.Format("destination={0}&amp;flags=0&amp;forcedownlevel=0&amp;trusted=0&amp;username={1}&amp;password={2}&amp;isUtf8=1", url, userName, password);
            string postContent = string.Format("destination={0}&flags=0&forcedownlevel=0&trusted=0&username={1}&password={2}&isUtf8=1;", url, userName, password);
            byte[] bytes = Encoding.UTF8.GetBytes(postContent);
            request.ContentLength = bytes.Length;

            // Create the web request content stream:

            using (Stream stream = request.GetRequestStream())
            {
                stream.Write(bytes, 0, bytes.Length);
                stream.Flush();
                stream.Close();
            }

            // Get the response & store the authentication cookies:
            HttpWebResponse response;
            try
            {
                response = (HttpWebResponse)request.GetResponse();
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Message: Connection failed to url ({0}). {1}", postUrl, ex.Message));
            }

            if (response.StatusCode != HttpStatusCode.OK)
            {

                throw new Exception(string.Format("Message: Connection failed to url ({0}). Response status code: '{1}'.", url, response.StatusCode));
            }

            using (Stream stream = response.GetResponseStream())
            {
                using (StreamReader sr = new StreamReader(stream))
                {
                    string result = sr.ReadToEnd();
                    sr.Close();
                }
                stream.Close();
            }

            if (response.Cookies.Count < 2)
                throw new AuthenticationException("Login failed. Is the login / password correct?");

            CookieContainer resultCookieContainer  = new CookieContainer();
            foreach (Cookie cookie in response.Cookies)
            {
                resultCookieContainer.Add(cookie);
            }

            response.Close();

            return resultCookieContainer;
        }

        /// 
        /// Add code here to check the server certificate, if you think it's necessary. Note that this
        /// will only be called once when the application is first started.
        /// 

        protected bool CheckCert(Object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        /// 
        /// Find today's appointments in the public folder calendar and print the results.
        /// 

//        public void RunQuery()
//        {
//            string uri = _hostName + path;
//            HttpWebRequest request;
//            WebResponse response;
//            byte[] bytes;

//            string start = DateTime.Today.ToString("yyyy/MM/dd");
//            string end = DateTime.Today.ToString("yyyy/MM/dd");

//            // Note that deep traversals don't work on public folders. In other words, if you
//            // need to dig deeper you'll need to split your query into multiple requests.

//            string format =
//                @"
//            
//                
//                    SELECT
//                        ""urn:schemas:calendar:dtstart"", ""urn:schemas:calendar:dtend"",
//                        ""urn:schemas:httpmail:subject"", ""urn:schemas:calendar:organizer"",
//                        ""DAV:parentname""
//                    FROM
//                        Scope('SHALLOW TRAVERSAL OF ""{0}""')
//                    WHERE
//                        NOT ""urn:schemas:calendar:instancetype"" = 1
//                        AND ""DAV:contentclass"" = 'urn:content-classes:appointment'
//                        AND ""urn:schemas:calendar:dtstart"" > '{1}'
//                        AND ""urn:schemas:calendar:dtend"" < '{2}'
//                    
//            ";

//            bytes = Encoding.UTF8.GetBytes(String.Format(format, uri, start, end));

//            // Use the authorization cookies we stored in the authentication method.

//            request = (HttpWebRequest)HttpWebRequest.Create(uri);
//            request.CookieContainer = _cookieContainerPost;
//            request.Method = "SEARCH";
//            request.ContentLength = bytes.Length;
//            request.ContentType = "text/xml";

//            using (Stream requestStream = request.GetRequestStream())
//            {
//                requestStream.Write(bytes, 0, bytes.Length);
//                requestStream.Close();
//            }

//            response = (HttpWebResponse)request.GetResponse();

//            using (Stream responseStream = response.GetResponseStream())
//            {
//                // Parse the XML response to find the data we need.

//                XmlDocument document = new XmlDocument();
//                document.Load(responseStream);

//                XmlNodeList subjectNodes = document.GetElementsByTagName("e:subject");
//                XmlNodeList locationNodes = document.GetElementsByTagName("a:parentname");
//                XmlNodeList startTimeNodes = document.GetElementsByTagName("d:dtstart");
//                XmlNodeList endTimeNodes = document.GetElementsByTagName("d:dtend");
//                XmlNodeList organizerNodes = document.GetElementsByTagName("d:organizer");

//                for (int index = 0; index < subjectNodes.Count; index++)
//                {
//                    string subject = subjectNodes[index].InnerText;
//                    string organizer = organizerNodes[index].InnerText;
//                    string location = /*ParentName(*/locationNodes[index].InnerText/*)*/;
//                    DateTime startTime = DateTime.Parse(startTimeNodes[index].InnerText);
//                    DateTime endTime = DateTime.Parse(endTimeNodes[index].InnerText);

//                    // Use a regex to get just the user's first and last names. Note that
//                    // some appointments may not have a valid user name.

//                    string pattern = @"""(?.*?)""";
//                    Regex regex = new Regex(pattern, RegexOptions.None);
//                    Match matchedText = regex.Match(organizer);

//                    if (matchedText.Success && matchedText.Groups["name"] != null)
//                        organizer = matchedText.Groups["name"].Value;

//                    // Print the results to the console.

//                    Console.WriteLine("{0} - {1}: {2}", startTime.ToShortTimeString(), endTime.ToShortTimeString(), subject);
//                    Console.WriteLine("{0} ({1})", location, organizer);
//                    Console.WriteLine();
//                }

//            }

//            response.Close();
//        }
    }
}
