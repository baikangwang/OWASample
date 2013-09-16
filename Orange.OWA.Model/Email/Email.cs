using System;
using Orange.OWA.Interface;

namespace Orange.OWA.Model.Email
{
    public class Email:IEmail
    {
        public string Id { get; set; }
        public string From { get; set; }
        public string Cc { get; set; }
        public string To { get; set; }
        public string Topic { get; set; }
        public string Subject { get; set; }
        public DateTime DateRecieved { get; set; }
        //public DateTime Date { get; set; }
        public string TextDescription { get; set; }
        public string HtmlDescription { get; set; }
        public bool Submitted { get; set; }
        public bool HasAttachment { get; set; }
        public bool Read { get; set; }
        public int Priority { get; set; }
        public string Url { get; set; }

    }
}
