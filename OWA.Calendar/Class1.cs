using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OWA.Calendar
{
    public class Class1
    {
        //<?xml version="1.0"?>
        //<searchrequest xmlns="DAV:">
        //<sql>SELECT "DAV:href" as URL, 
        //"http://schemas.microsoft.com/exchange/outlookmessageclass" as Class, 
        //"urn:schemas:calendar:dtstart" as StartTime, 
        //"urn:schemas:calendar:dtend" as EndTime, 
        //"urn:schemas:calendar:remindernexttime" as ReminderTime, 
        //"urn:schemas:calendar:location" as Location, 
        //"http://schemas.microsoft.com/mapi/subject" as Subject, 
        //"urn:schemas:calendar:instancetype" as Type, 
        //"http://schemas.microsoft.com/exchange/smallicon" as Icon 
        // from SCOPE('shallow traversal of ""') 
        // WHERE NOT ("urn:schemas:calendar:instancetype" = 1) 
        // AND ("urn:schemas:calendar:remindernexttime" &lt; CAST("2013-09-10T16:00:00Z" as "dateTime.tz")) 
        // AND ("urn:schemas:calendar:dtstart" &lt; CAST("2013-12-08T16:00:00Z" as "dateTime.tz")) 
        // AND ("urn:schemas:calendar:dtend"   &gt; CAST("2013-06-08T16:00:00Z" as "dateTime.tz")) 
        // AND ("DAV:ishidden" is Null  OR "DAV:ishidden" = false) 
        //ORDER BY "urn:schemas:calendar:dtstart" DESC </sql>
        //</searchrequest>    
    }
}
