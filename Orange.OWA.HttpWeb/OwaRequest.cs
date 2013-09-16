using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

namespace Orange.OWA.HttpWeb
{
    public class OwaRequest
    {
        private readonly HttpWebRequest _request;
        private readonly byte[] _content;

        protected OwaRequest(string method, string url, params object[] args)
        {
            _request = (HttpWebRequest)System.Net.WebRequest.Create(url);
            _request.Method = method;
            _request.Accept = "text/html, application/xhtml+xml, */*";
            _request.Headers.Add("Accept-Language", "en-US");
            _request.Headers.Add("Accept-Encoding", "gzip,deflate");
            _request.UserAgent = "Orange.OWA";
            _request.ContentType = "application/x-www-form-urlencoded";
            _request.KeepAlive = true;
            _request.Headers.Add("Cache-Control", "no-cache");

            if (args == null)
                return;

            if (args.Length > 0)
            {
                byte[] content = args[0] as byte[];
                if (content != null)
                {
                    _request.ContentLength = content.Length;
                    this._content = content;
                }
            }
            
            if (args.Length > 1)
            {
                IList<Cookie> cookies = args[1] as IList<Cookie>;
                CookieContainer cookieContainer=new CookieContainer();
                if (cookies != null)
                {
                    foreach (Cookie cookie in cookies)
                    {
                        cookieContainer.Add(cookie);
                    }
                }

                _request.CookieContainer = cookieContainer;
            }

            
            if (args.Length > 2)
            {
                IDictionary<string, string> headers = args[2] as IDictionary<string, string>;
                if (headers != null)
                {
                    foreach (string key in headers.Keys)
                    {
                        _request.Headers.Add(key, headers[key]);
                    }
                }
            }
        }

        public OwaResponse Send()
        {
            if (_content != null)
            {
                using (Stream s = _request.GetRequestStream())
                {
                    s.Write(_content, 0, _content.Length);
                    s.Flush();
                }
            }

            HttpWebResponse response;
            try
            {
                response = (HttpWebResponse)_request.GetResponse();
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Message: request failed to url ({0}). {1}", _request.RequestUri, ex.Message));
            }

            OwaResponse owaOwaResponse=new OwaResponse(response);

            return owaOwaResponse;
        }

        public string Accept
        {
            get { return _request.Accept; }
            set { _request.Accept = value; }
        }

        public string ContentType
        {
            get { return _request.ContentType; }
            set { _request.ContentType = value; }
        }

        public static OwaRequest Get(string url)
        {
            return new OwaRequest("Get",url);
        }

        public static OwaRequest Post(string url, byte[] content, IList<Cookie> cookies = null, IDictionary<string, string> headers = null)
        {
            return new OwaRequest("POST", url, content, cookies, headers);
        }

        public static OwaRequest Search(string url, byte[] content, IList<Cookie> cookies = null, IDictionary<string, string> headers = null)
        {
            return new OwaRequest("SEARCH", url, content, cookies, headers);
        }
    }
}
