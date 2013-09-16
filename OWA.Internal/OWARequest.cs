using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Cache;
using System.Security.Authentication;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Xml;

namespace OWA.Internal
{
    public class OWARequest
    {
        private CookieContainer _cookieContainerPost;
        private string _hostName = "webmail.taylorcorp.com";
        private string _url = "https://webmail.taylorcorp.com/exchange";
        private string _postUrl = "https://webmail.taylorcorp.com/exchweb/bin/auth/owaauth.dll";
        private string calendarPath = "https://webmail.taylorcorp.com/exchange/bkwang@nltechdev.com/Calendar";
        private string _username = "corp\\bkwang";
        private string _password = "R8ll#qqO2";
        //private int _requestTimeout = 15000;
        
        public void Authenticate()
        {
            _cookieContainerPost = OWAPOST(this._hostName, this._url, this._postUrl, this._username, this._password/*, get*/);
        }

        private CookieContainer OWAPOST(string hostName, string url, string postUrl, string userName, string password, CookieContainer getCookieContainer = null)
        {
            // Create the web request:

            HttpWebRequest request = (HttpWebRequest)System.Net.WebRequest.Create(postUrl);
            request.Method = "POST";
            request.Accept = "text/html, application/xhtml+xml, */*";
            request.Headers.Add("Accept-Language", "en-US");
            request.Headers.Add("Accept-Encoding", "gzip,deflate");
            request.UserAgent = "OWASample";
            request.ContentType = "application/x-www-form-urlencoded";
            request.KeepAlive = true;
            request.Headers.Add("Cache-Control", "no-cache");
            request.CookieContainer = new CookieContainer();

            string postContent =  string.Format("destination={0}&flags=0&forcedownlevel=0&trusted=0&username={1}&password={2}&SubmitCreds=Log On;", url, userName, password);
            byte[] bytes = Encoding.UTF8.GetBytes(postContent);
            request.ContentLength = bytes.Length;

            // Create the web request content stream:
            using (Stream s = request.GetRequestStream())
            {
                s.Write(bytes, 0, bytes.Length);
                s.Flush();
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

            if (response.Cookies.Count < 2)
                throw new AuthenticationException("Login failed. Is the login / password correct?");

            CookieContainer resultCookieContainer = new CookieContainer();
            foreach (Cookie cookie in response.Cookies)
            {
                resultCookieContainer.Add(cookie);
            }

            response.Close();

            return resultCookieContainer;
        }

        public void RunQuery()
        {
            string uri = calendarPath; //"https://webmail.taylorcorp.com/exchange/bkwang@nltechdev.com/Inbox/";//calendarPath;
            HttpWebRequest request;
            HttpWebResponse response;

            byte[] bytes;

            string start = DateTime.Today.ToString("yyyy/MM/dd");
            string end = DateTime.Today.ToString("yyyy/MM/dd");

            // Note that deep traversals don't work on public folders. In other words, if you
            // need to dig deeper you'll need to split your query into multiple requests.
            StringBuilder sb=new StringBuilder();

            //sb.AppendLine("<?xml version=\"1.0\"?>");
            //sb.AppendLine("<searchrequest xmlns=\"DAV:\">");
            //sb.AppendLine("	<sql>SELECT \"http://schemas.microsoft.com/exchange/smallicon\" as smicon, \"http://schemas.microsoft.com/mapi/sent_representing_name\" as from, \"urn:schemas:httpmail:datereceived\" as recvd, \"http://schemas.microsoft.com/mapi/proptag/x10900003\" as flag, \"http://schemas.microsoft.com/mapi/subject\" as subj, \"http://schemas.microsoft.com/exchange/x-priority-long\" as prio, \"urn:schemas:httpmail:hasattachment\" as fattach,\"urn:schemas:httpmail:read\" as r, \"http://schemas.microsoft.com/exchange/outlookmessageclass\" as m, \"http://schemas.microsoft.com/mapi/proptag/x10950003\" as flagcolor");
            //sb.AppendLine("FROM Scope('SHALLOW TRAVERSAL OF \"\"')");
            //sb.AppendLine("WHERE \"http://schemas.microsoft.com/mapi/proptag/0x67aa000b\" = false AND \"DAV:isfolder\" = false");
            //sb.AppendLine("ORDER BY \"urn:schemas:httpmail:datereceived\" DESC");
            //sb.AppendLine("	</sql>");
            //sb.AppendLine("	<range type=\"row\">0-24</range>");
            //sb.AppendLine("</searchrequest>");          
            sb.AppendLine("<?xml version=\"1.0\"?>");
            sb.AppendLine("<searchrequest xmlns=\"DAV:\">");
            sb.AppendLine("<sql>SELECT \"DAV:href\" as URL,");
            sb.AppendLine("\"http://schemas.microsoft.com/exchange/outlookmessageclass\" as Class,");
            sb.AppendLine("\"urn:schemas:calendar:dtstart\" as StartTime,");
            sb.AppendLine("\"urn:schemas:calendar:dtend\" as EndTime,");
            sb.AppendLine("\"urn:schemas:calendar:remindernexttime\" as ReminderTime,");
            sb.AppendLine("\"urn:schemas:calendar:location\" as Location,");
            sb.AppendLine("\"http://schemas.microsoft.com/mapi/subject\" as Subject,");
            sb.AppendLine("\"urn:schemas:calendar:instancetype\" as Type, ");
            sb.AppendLine("\"http://schemas.microsoft.com/exchange/smallicon\" as Icon");
            sb.AppendLine(" from SCOPE('shallow traversal of \"\"')");
            sb.AppendLine(" WHERE NOT (\"urn:schemas:calendar:instancetype\" = 1)");
            sb.AppendLine(" AND (\"urn:schemas:calendar:remindernexttime\" &lt; CAST(\"2013-09-10T16:00:00Z\" as \"dateTime.tz\"))");
            sb.AppendLine(" AND (\"urn:schemas:calendar:dtstart\" &lt; CAST(\"2013-12-08T16:00:00Z\" as \"dateTime.tz\"))");
            sb.AppendLine(" AND (\"urn:schemas:calendar:dtend\" &gt; CAST(\"2013-06-08T16:00:00Z\" as \"dateTime.tz\"))");
            sb.AppendLine(" AND (\"DAV:ishidden\" is Null  OR \"DAV:ishidden\" = false)");
            sb.AppendLine("ORDER BY \"urn:schemas:calendar:dtstart\" DESC </sql>");
            sb.AppendLine("</searchrequest>");
            string format = sb.ToString();
            bytes = Encoding.UTF8.GetBytes(format/*String.Format(format, uri, start, end)*/);

            // Use the authorization cookies we stored in the authentication method.
            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3;
            request = (HttpWebRequest) System.Net.WebRequest.Create(uri);
            request.Method = "SEARCH";
            request.Accept = "*/*";
            request.Headers.Add("depth", "1");//calendar
            request.Headers.Add("Translate", "f");
            //request.Headers.Add("brief", "t");//inbox
            request.ContentType = "text/xml";
            request.Headers.Add("Accept-Language", "en-US");
            request.Headers.Add("Accept-Encoding", "gzip,deflate");
            request.UserAgent = "OWASample";
            //request.Host = _hostName;
            request.ContentLength = bytes.Length;
            request.KeepAlive = true;
            request.Headers.Add("Cache-Control", "no-cache");
            request.CookieContainer = _cookieContainerPost;
            //request.AllowAutoRedirect = true;

            using (Stream requestStream = request.GetRequestStream())
            {
                requestStream.Write(bytes, 0, bytes.Length);
                requestStream.Flush();
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
