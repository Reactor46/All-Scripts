using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace PSReport
{
    public class Server
    {
        [XmlText]
        public string Name { get; set; }
        [XmlAttribute]
        public string Type { get; set; }

        public Server()
        {

        }

        public Server(string Name, string Type)
        {
            this.Name = Name;
            this.Type = Type;
        }

    }
}
