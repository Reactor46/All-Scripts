using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Inventory.Model
{
    public class AppPool
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string State { get; set; }
        public string Username { get; set; }
        public int IISVersion { get; set; }
        public DateTime Timestamp { get; set; }
    }
}
