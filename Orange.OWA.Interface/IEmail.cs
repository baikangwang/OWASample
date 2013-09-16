using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Orange.OWA.Interface
{
    public interface IEmail
    {
        string Id { get; set; }
        string From { get; set; }                   //"Ding, Li (Northern Lights)" &lt;lding@nltechdev.com&gt;
        //string FromName { get; set; }               //Ding, Li (Northern Lights)
        //string FromMail { get; set; }             //lding@nltechdev.com
        string Cc { get; set; }                     //"Tran, Lam D (PaperDirect)" &lt;LDTran@PaperDirect.com&gt;, "Zhang, Daniel (Current)" &lt;DZhang@currentinc.com&gt;
        //string DisplayCc { get; set; }
        string To { get; set; }
        //string DisplayTo { get; set; }              //"Welch, Mike J (Current); Wang, Bai Kang (Northern Lights)"
        string Topic { get; set; }                  //CMFG-431 / Universal converter concern
        string Subject { get; set; }
        //string NormalizedSubject { get; set; }    //CMFG-431 / Universal converter concern
        DateTime DateRecieved { get; set; }         //dateTime.tz:2013-09-10T03:39:46.190Z
        //DateTime Date { get; set; }                 //dateTime.tz:2013-09-10T03:39:41.143Z
        //string SenderName { get; set; }           //Ding, Li (Northern Lights)         
        string TextDescription { get; set; }
        string HtmlDescription { get; set; }
        bool Submitted { get; set; }
        bool HasAttachment { get; set; }
        bool Read { get; set; }
        int Priority { get; set; }
        string Url { get; set; }
    }
}
