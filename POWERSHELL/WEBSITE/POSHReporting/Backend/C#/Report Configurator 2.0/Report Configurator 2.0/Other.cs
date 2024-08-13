using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace PSReport
{
    public class Other
    {
        [XmlAttribute]
        public Boolean Enabled { get; set; }
        public Script Script { get; set; }

        public Other()
        {

        }
        public Other(Boolean enabled, Script script)
        {
            Enabled = enabled;
            Script = script;
        }

    }
}
