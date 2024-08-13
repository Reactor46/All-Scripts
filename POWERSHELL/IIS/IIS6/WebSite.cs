using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Inventory.Client.IIS6
{
    public class Website
    {
        public Int32 Identity
        {
            get;
            set;
        }

        public String Name
        {
            get;
            set;
        }

        public String PhysicalPath
        {
            get;
            set;
        }

        public ServerState Status
        {
            get;
            set;
        }
    }
}
