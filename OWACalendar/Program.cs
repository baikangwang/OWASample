using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Security;
using System.Security.Authentication;
using System.Text;
using OWASample;

namespace OWACalendar
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                WebDAVRequest request = new WebDAVRequest();
                request.Authenticate();
                //request.RunQuery();
            }
            catch (SecurityException e)
            {
                Console.WriteLine("Security Exception");
                Console.WriteLine("   Msg: " + e.Message);
                Console.WriteLine("   Note: The application may not be trusted if run from a network share.");
            }
            catch (AuthenticationException e)
            {
                Console.WriteLine("Authentication Exception, are you using a valid login?");
                Console.WriteLine("   Msg: " + e.Message);
                Console.WriteLine("   Note: You must use a valid login / password for authentication.");
            }
            catch (WebException e)
            {
                Console.WriteLine("Web Exception");
                Console.WriteLine("   Status: " + e.Status);
                Console.WriteLine("   Reponse: " + e.Response);
                Console.WriteLine("   Msg: " + e.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine("Unknown Exception");
                Console.WriteLine("   Msg: " + e.Message);
            }

            Console.Read();
        }
    }
}
