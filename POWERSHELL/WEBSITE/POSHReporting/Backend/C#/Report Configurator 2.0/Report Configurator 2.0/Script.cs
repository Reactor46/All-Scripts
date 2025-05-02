using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace PSReport
{
    public class Script
    {
        [XmlText]
        public string Title { get; set; }
        [XmlAttribute]
        public string Path { get; set; }
        [XmlAttribute]
        public Boolean Enabled { get; set; }

        public Script()
        {

        }

        public Script(string title, string path, Boolean enabled)
        {
            Title = title;
            Path = path;
            Enabled = enabled;
        }
    }
}
