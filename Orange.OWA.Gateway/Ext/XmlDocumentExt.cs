using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Orange.OWA.Gateway
{
    public static class XmlDocumentExt
    {
        public static T GetNodeValueByKey<T>(this XmlNode node, string key, XmlNamespaceManager nsmgr)
        {
            string path = string.Format("a:propstat/a:prop/{0}", key);
            return GetNodeValue<T>(node, path, nsmgr);
        }

        public static T GetNodeValue<T>(this XmlNode node, string path, XmlNamespaceManager nsmgr)
        {
            XmlNode target = node.SelectSingleNode(path, nsmgr);
            if (target == null || string.IsNullOrEmpty(target.InnerText))
                return default(T);

            object value = null;
            Type t = typeof(T);
            if (t == typeof(string))
            {
                value = target.InnerText;
            }
            else if (t == typeof(DateTime))
            {
                value = Convert.ToDateTime(target.InnerText);//DateTime.ParseExact(target.InnerText, "yyyy-MM-ddThh:mm:ss.fffZ", System.Globalization.CultureInfo.InvariantCulture);// 
            }
            else if (t == typeof(bool))
            {
                value = Convert.ToBoolean(Convert.ToInt32(target.InnerText));
            }
            else if (t == typeof(int))
            {
                value = Convert.ToInt32(target.InnerText);
            }

            return value == null ? default(T) : (T)Convert.ChangeType(value, typeof(T));
        }
    }
}
