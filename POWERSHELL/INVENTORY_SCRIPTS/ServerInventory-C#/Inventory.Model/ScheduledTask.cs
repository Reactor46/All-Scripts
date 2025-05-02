using System;

namespace Inventory.Model
{
    public class ScheduledTask
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public bool Enabled { get; set; }
        public string Principal { get; set; }
        public DateTime Timestamp { get; set; }
    }
}