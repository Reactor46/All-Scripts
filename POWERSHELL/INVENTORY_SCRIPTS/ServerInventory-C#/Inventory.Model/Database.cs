using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Inventory.Model
{
    public class Database
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public bool Online { get; set; }
        public DateTime Timestamp { get; set; }

    }
}
