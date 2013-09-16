using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OWA.Task
{
    public class Class1
    {
        //<?xml version="1.0"?>
        //<searchrequest xmlns="DAV:">
        //<sql>SELECT "DAV:href" as URL, 
        //"http://schemas.microsoft.com/exchange/outlookmessageclass" as Class, 
        //"http://schemas.microsoft.com/mapi/commonend" as StartTime, 
        //"http://schemas.microsoft.com/mapi/remindernexttime" as ReminderTime, 
        //"http://schemas.microsoft.com/mapi/subject" as Subject, 
        //"http://schemas.microsoft.com/exchange/smallicon" as Icon 
        // from SCOPE('shallow traversal of ""') 
        // WHERE ("http://schemas.microsoft.com/exchange/tasks/is_complete" = false ) 
        // AND ("http://schemas.microsoft.com/mapi/reminderset" = true) 
        // AND ("http://schemas.microsoft.com/mapi/remindernexttime" &lt; CAST("2013-09-10T16:00:00Z" as "dateTime.tz")) 
        // AND (("http://schemas.microsoft.com/mapi/commonend" IS NULL) OR (("http://schemas.microsoft.com/mapi/commonend" &lt; CAST("2013-12-08T16:00:00Z" as "dateTime.tz")) 
        // AND ("http://schemas.microsoft.com/mapi/commonend" &gt; CAST("2013-06-08T16:00:00Z" as "dateTime.tz")))) 
        // AND ("DAV:ishidden" = false) 
        // ORDER BY "http://schemas.microsoft.com/exchange/tasks/dtdue" DESC </sql>
        //</searchrequest>    
    }
}
