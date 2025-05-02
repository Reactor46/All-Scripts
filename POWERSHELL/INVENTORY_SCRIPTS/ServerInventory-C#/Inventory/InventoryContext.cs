using Inventory.Model;
using System.Data.Entity;


namespace Inventory.Server
{
    public class InventoryContext:DbContext
    {

        public InventoryContext()
        {
            // remove this line if you want a new database to be created locally automatically 
            this.Database.Connection.ConnectionString = Properties.Settings.Default.InventoryDb;
        }
      
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder
                .Entity<Model.Server>()
                .HasMany(x => x.Databases)
                .WithOptional()
                .WillCascadeOnDelete();
            modelBuilder
                .Entity<Model.Server>()
                .HasMany(x => x.DatabaseJobs)
                .WithOptional()
                .WillCascadeOnDelete();
            modelBuilder
                .Entity<Model.Server>()
                .HasMany(x => x.VirtualDirectories)
                .WithOptional()
                .WillCascadeOnDelete();
            modelBuilder
                .Entity<Model.Server>()
                .HasMany(x => x.ScheduledTasks)
                .WithOptional()
                .WillCascadeOnDelete();
            modelBuilder
                .Entity<Model.Server>()
                .HasMany(x => x.ApplicationPools)
                .WithOptional()
                .WillCascadeOnDelete();


        }

    
        public DbSet<Model.Server> Servers { get; set; }
        public DbSet<Model.Database> Databases { get; set; }
        public DbSet<DatabaseJob> DatabaseJobs { get; set; }
        public DbSet<ScheduledTask> ScheduledTasks { get; set; }
        public DbSet<VirtualDir> VirtualDirectories { get; set; }
        public DbSet<AppPool> ApplicationPools { get; set; }
    }
}