using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace PSReport
{
    public class Configuration
    {
        public General General { get; set; }
        public Email Email { get; set; }
        public ObservableCollection<Script> Scripts { get; set; }
        public ObservableCollection<Server> Servers { get; set; }
        public Other Other { get; set; }

        public Configuration()
        {

        }

        public Configuration(General general, Email email, ObservableCollection<Script> scripts, ObservableCollection<Server> servers, Other other)
        {
            General = general;
            Email = email;
            Servers = servers;
            Scripts = scripts;
            Other = other;
        }

        public void Save(string filePath)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Configuration));

            using (TextWriter tw = new StreamWriter(filePath))
            {
                xmlSerializer.Serialize(tw, this);
            }

        }

        public static Configuration Load(string filePath)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Configuration));

            using (FileStream Stream = new FileStream(filePath, FileMode.Open))
            {
                XmlReader reader = XmlReader.Create(Stream);

                return (Configuration)xmlSerializer.Deserialize(reader);
            }
        }
    }
}
