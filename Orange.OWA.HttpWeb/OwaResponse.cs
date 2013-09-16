using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

namespace Orange.OWA.HttpWeb
{
    public class OwaResponse:IDisposable
    {
        private HttpWebResponse _response;
        private IList<Cookie> _cookies;

        internal OwaResponse(HttpWebResponse response)
        {
            _response = response;

            _cookies = new List<Cookie>();
            
            if (response.Cookies != null)
            {
                foreach (Cookie cookie in response.Cookies)
                {
                    _cookies.Add(cookie);
                }
            }
        }

        public IList<Cookie> Cookies { get { return _cookies; } }

        public Stream GetResponseStream()
        {
            return _response.GetResponseStream();
        }

        public string GetResponseHeader(string key)
        {
            return _response.GetResponseHeader(key);
        }

        public HttpStatusCode StatusCode
        {
            get { return _response.StatusCode; }
        }

        public void Close()
        {
            _response.Close();
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool cleanupAll)
        {
            if (cleanupAll)
            {
                Close();
                Dispose(false);
            }
            else
            {
                _cookies = null;
            }
        }
    }
}
