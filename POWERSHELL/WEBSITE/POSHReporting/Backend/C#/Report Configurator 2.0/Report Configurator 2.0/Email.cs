using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Xml.Serialization;

namespace PSReport
{
    public class Email
    {
        [XmlAttribute]
        public Boolean Enabled { get; set; }

        public string SMTPServer { get; set; }
        public int Port { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string Cc { get; set; }
        public string Bcc { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }

        public Email()
        {

        }

        public Email(Boolean enabled, string smptserver, int port, string from, string to, string cc, string bcc, string subject, string body)
        {
            Enabled = enabled;
            SMTPServer = smptserver;
            Port = port;
            From = from;
            To = to;
            Cc = cc;
            Bcc = bcc;
            Subject = subject;
            Body = body;
        }

    }
}