using System;

namespace Inventory.Model
{
    public class VirtualDir
    {

        public int Id { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public string State { get; set; }
        public string ApplicationPool { get; set; }
        public int IISVersion { get; set; }
        public DateTime Timestamp { get; set; }
    }
}