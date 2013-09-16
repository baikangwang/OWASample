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
        private string _url = "https://legacymail.taylorcorp.com/exchange";//"https://webmail.taylorcorp.com/owa";
        private string _postUrl = "https://legacymail.taylorcorp.com/exchweb/bin/auth/owaauth.dll";//"https://legacymail.taylorcorp.com/exchweb/bin/auth/owaauth.dll";//"https://webmail.taylorcorp.com/owa/auth.owa";
        //private string path = "/bkwang@nltechdev.com/Calendar";//"/public/Calendars/Board%20Room";
        private string calendarPath = "https://legacymail.taylorcorp.com/exchange/bkwang@nltechdev.com/Calendar";
        private  string _username = "corp\\bkwang";//"user";
        private  string _password = "R8ll#qqO2";//"password";
        private CookieContainer _cookieContainerPost = null;
        //private CookieContainer _cookieContainerGet = null;
        private int _requestTimeout = 15000;
        //private string _successLoginText = "New Message";
       // private TimeSpan _responseTime=new TimeSpan();


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
            //CookieContainer get = OWAGET(this._url);
            _cookieContainerPost = OWAPOST(this._hostName, this._url, this._postUrl,this._username,this._password/*, get*/);
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

            Stream stream = response.GetResponseStream();
                if (stream != null)
                    using (StreamReader sr = new StreamReader(stream))
                    {
                        string result = sr.ReadToEnd();
                    }

            CookieContainer getCookieContainer = new CookieContainer();
            foreach (Cookie cookie in response.Cookies)
            {
                getCookieContainer.Add(cookie);
            }

            response.Close();

            return getCookieContainer;
        }

        private CookieContainer OWAPOST(string hostName, string url,string postUrl,string userName,string password, CookieContainer getCookieContainer=null)
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

            //Uri uri = new Uri(_url);
            //CookieContainer postCookieContainer = new CookieContainer();
            //Cookie outlookJavascriptCookie = null;
            //if (getCookieContainer.Count > 0)
            //{
            //    Cookie outlookSessionCookie = getCookieContainer.GetCookies(uri)[0];
            //    outlookJavascriptCookie=new Cookie("PBack","0",outlookSessionCookie.Path,outlookSessionCookie.Domain);
            //}
            //else
            //{
            //    outlookJavascriptCookie = new Cookie("PBack", "0", "/", hostName);
            //}
            //postCookieContainer.Add(outlookJavascriptCookie);
            //request.CookieContainer = postCookieContainer;
            request.CookieContainer=new CookieContainer();

            //string postContent = string.Format("destination={0}&amp;flags=0&amp;forcedownlevel=0&amp;trusted=0&amp;username={1}&amp;password={2}&amp;isUtf8=1", url, userName, password);
            string postContent = string.Format("destination={0}&flags=0&forcedownlevel=0&trusted=0&username={1}&password={2}&isUtf8=1;", url, userName, password);
            byte[] bytes = Encoding.UTF8.GetBytes(postContent);
            request.ContentLength = bytes.Length;

            // Create the web request content stream:

            using (Stream s = request.GetRequestStream())
            {
                s.Write(bytes, 0, bytes.Length);
                s.Flush();
                //stream.Close();
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

            Stream stream = response.GetResponseStream();
                if (stream != null)
                    using (StreamReader sr = new StreamReader(stream))
                    {
                        string result = sr.ReadToEnd();
                    }

            if (response.Cookies.Count < 2)
                throw new AuthenticationException("Login failed. Is the login / password correct?");

            CookieContainer resultCookieContainer = new CookieContainer();
            foreach (Cookie cookie in response.Cookies)
            {
                resultCookieContainer.Add(cookie);
            }

            response.Close();

            return resultCookieContainer;
            //return new CookieContainer();
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

        public void RunQuery()
        {
            string uri = calendarPath;//_url + path;
            HttpWebRequest request;
            WebResponse response;
            byte[] bytes;

            string start = DateTime.Today.ToString("yyyy/MM/dd");
            string end = DateTime.Today.ToString("yyyy/MM/dd");

            // Note that deep traversals don't work on public folders. In other words, if you
            // need to dig deeper you'll need to split your query into multiple requests.
//            <a:searchrequest xmlns:a="DAV:" > 
//<a:sql>Select "urn:schemas:calendar:dtstart" AS dtstart, 
//"urn:schemas:calendar:dtend" AS dtend, 
//"urn:schemas:calendar:alldayevent" AS alldayevent 
//FROM Scope('SHALLOW TRAVERSAL OF ""') 
//WHERE NOT "urn:schemas:calendar:instancetype" = 1 AND "urn:schemas:calendar:dtend" &gt; '2013/08/24 16:00:00' 
//AND "urn:schemas:calendar:dtstart" &lt; '2013/10/05 16:00:00' 
//AND NOT "urn:schemas:calendar:instancetype" = 1 
//AND NOT "urn:schemas:calendar:busystatus" = 'FREE' 
//ORDER BY "urn:schemas:calendar:dtstart" ASC 
//</></>
//            string format =
//                @"
//                    SELECT
//                        ""urn:schemas:calendar:dtstart"" As dtstart, ""urn:schemas:calendar:dtend"" As dtend,
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

            string format =
            "<?xml version=\"1.0\"?>"+
            "<searchrequest xmlns=\"DAV:\">"+
            "<sql>SELECT \"DAV:href\" as URL,"+ 
            "\"http://schemas.microsoft.com/exchange/outlookmessageclass\" as Class, "+ 
            "\"urn:schemas:calendar:dtstart\" as StartTime,"+ 
            "\"urn:schemas:calendar:dtend\" as EndTime,"+ 
            "\"urn:schemas:calendar:remindernexttime\" as ReminderTime, "+ 
            "\"urn:schemas:calendar:location\" as Location,"+ 
            "\"http://schemas.microsoft.com/mapi/subject\" as Subject, "+ 
            "\"urn:schemas:calendar:instancetype\" as Type, "+
            "\"http://schemas.microsoft.com/exchange/smallicon\" as Icon "+
            " from SCOPE('shallow traversal of \"\"') "+
            " WHERE NOT (\"urn:schemas:calendar:instancetype\" = 1) "+
            " AND (\"urn:schemas:calendar:remindernexttime\" < CAST(\"2013-09-10T16:00:00Z\" as \"dateTime.tz\")) "+
            " AND (\"urn:schemas:calendar:dtstart\" < CAST(\"2013-12-08T16:00:00Z\" as \"dateTime.tz\")) "+
            " AND (\"urn:schemas:calendar:dtend\"   > CAST(\"2013-06-08T16:00:00Z\" as \"dateTime.tz\")) "+
            " AND (\"DAV:ishidden\" is Null  OR \"DAV:ishidden\" = false) "+
            "ORDER BY \"urn:schemas:calendar:dtstart\" DESC </sql>"+
            "</searchrequest>";
            //string format =
            //"<?xml version=\"1.0\"?>" +
            //"<searchrequest xmlns=\"DAV:\">" +
            //"<sql>SELECT \"DAV:href\" as URL," +
            //"\"http://schemas.microsoft.com/exchange/outlookmessageclass\" as Class, " +
            //"\"urn:schemas:calendar:dtstart\" as StartTime," +
            //"\"urn:schemas:calendar:dtend\" as EndTime," +
            //"\"urn:schemas:calendar:remindernexttime\" as ReminderTime, " +
            //"\"urn:schemas:calendar:location\" as Location," +
            //"\"http://schemas.microsoft.com/mapi/subject\" as Subject, " +
            //"\"urn:schemas:calendar:instancetype\" as Type, " +
            //"\"http://schemas.microsoft.com/exchange/smallicon\" as Icon " +
            //" from SCOPE('shallow traversal of \"\"') " +
            //" WHERE NOT (\"urn:schemas:calendar:instancetype\" = 1) " +
            //" AND (\"urn:schemas:calendar:remindernexttime\" &lt; CAST(\"2013-09-10T16:00:00Z\" as \"dateTime.tz\")) " +
            //" AND (\"urn:schemas:calendar:dtstart\" &lt; CAST(\"2013-12-08T16:00:00Z\" as \"dateTime.tz\")) " +
            //" AND (\"urn:schemas:calendar:dtend\"   &gt; CAST(\"2013-06-08T16:00:00Z\" as \"dateTime.tz\")) " +
            //" AND (\"DAV:ishidden\" is Null  OR \"DAV:ishidden\" = false) " +
            //"ORDER BY \"urn:schemas:calendar:dtstart\" DESC </sql>" +
            //"</searchrequest>";
            bytes = Encoding.UTF8.GetBytes(format/*String.Format(format, uri, start, end)*/);

            // Use the authorization cookies we stored in the authentication method.
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3;
            request = (HttpWebRequest)HttpWebRequest.Create(uri);
            request.CookieContainer = _cookieContainerPost;
            request.Method = "SEARCH";
            request.ContentLength = bytes.Length;
            request.ContentType = "text/xml";
            

            using (Stream requestStream = request.GetRequestStream())
            {
                requestStream.Write(bytes, 0, bytes.Length);
                //requestStream.Flush();
                //requestStream.Close();
            }

            response = (HttpWebResponse)request.GetResponse();

            using (Stream responseStream = response.GetResponseStream())
            {
                // Parse the XML response to find the data we need.

                XmlDocument document = new XmlDocument();
                if (responseStream != null) document.Load(responseStream);

                XmlNodeList subjectNodes = document.GetElementsByTagName("e:subject");
                XmlNodeList locationNodes = document.GetElementsByTagName("a:parentname");
                XmlNodeList startTimeNodes = document.GetElementsByTagName("d:dtstart");
                XmlNodeList endTimeNodes = document.GetElementsByTagName("d:dtend");
                XmlNodeList organizerNodes = document.GetElementsByTagName("d:organizer");

                for (int index = 0; index < subjectNodes.Count; index++)
                {
                    string subject = subjectNodes[index].InnerText;
                    string organizer = organizerNodes[index].InnerText;
                    string location = /*ParentName(*/locationNodes[index].InnerText/*)*/;
                    DateTime startTime = DateTime.Parse(startTimeNodes[index].InnerText);
                    DateTime endTime = DateTime.Parse(endTimeNodes[index].InnerText);

                    // Use a regex to get just the user's first and last names. Note that
                    // some appointments may not have a valid user name.

                    string pattern = @"""(?.*?)""";
                    Regex regex = new Regex(pattern, RegexOptions.None);
                    Match matchedText = regex.Match(organizer);

                    if (matchedText.Success && matchedText.Groups["name"] != null)
                        organizer = matchedText.Groups["name"].Value;

                    // Print the results to the console.

                    Console.WriteLine("{0} - {1}: {2}", startTime.ToShortTimeString(), endTime.ToShortTimeString(), subject);
                    Console.WriteLine("{0} ({1})", location, organizer);
                    Console.WriteLine();
                }

            }

            response.Close();
        }
    }
}
