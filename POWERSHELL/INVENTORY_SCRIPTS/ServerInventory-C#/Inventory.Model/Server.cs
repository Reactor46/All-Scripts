using System;
using System.Collections.Generic;

namespace Inventory.Model
{
    public class Server
    {
        public Server()
        {
            Databases = new List<Database>();
            ScheduledTasks = new List<ScheduledTask>();
        }

        public string Id { get; set; }
        public virtual List<Database> Databases { get; set; }
        public virtual List<ScheduledTask> ScheduledTasks { get; set; }
        public string Domain { get; set; }
        public string ConnectionString { get; set; }
        public List<DatabaseJob> DatabaseJobs { get; set; }
        public List<VirtualDir> VirtualDirectories { get; set; }
        public List<AppPool> ApplicationPools { get; set; }
        public DateTime Timestamp { get; set; }
    }
}
