using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using KlerksSoft;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Orange.OWA.Test
{
    [TestClass]
    public class AuthenticationTest
    {
        [TestMethod]
        public void LogonTest()
        {
            IList<Cookie> cookies = Orange.OWA.Authentication.AuthenticationManager.Current.CookieCache;
            
            Assert.IsTrue(cookies!=null);
            
            Console.WriteLine("Count:{0}",cookies.Count);
            foreach (Cookie cookie in cookies)
            {
                Console.WriteLine("Index:{0}",cookies.IndexOf(cookie));
                Console.WriteLine("\tDomain:{0}", cookie.Domain);
                Console.WriteLine("\tName:{0}", cookie.Name);
                string value = cookie.Value;
                if (cookie.Name.ToLower() == "cadata")
                {
                    byte[] content = Convert.FromBase64String(value.Substring(1, value.Length - 2));
                    Encoding encoding = TextFileEncodingDetector.DetectTextByteArrayEncoding(content); //Encoding.GetEncoding("gzip");
                    if (encoding != null)
                    {
                        value = encoding.GetString(content);
                        Console.WriteLine("\tBase64 Value:{0}", cookie.Value);
                    }
                }
                Console.WriteLine("\tValue:{0}", value);
                Console.WriteLine("\tPort:{0}", string.IsNullOrEmpty(cookie.Port)?"N/A":cookie.Port);
                Console.WriteLine("\tPath:{0}", string.IsNullOrEmpty(cookie.Port) ? "N/A" : cookie.Path);
                Console.WriteLine("\tHttpOnly:{0}", cookie.HttpOnly);
                Console.WriteLine("\tComment:{0}", string.IsNullOrEmpty(cookie.Port) ? "N/A" : cookie.Comment);
            }

            Assert.IsTrue(cookies.Count!=0);
        }
    }
}
