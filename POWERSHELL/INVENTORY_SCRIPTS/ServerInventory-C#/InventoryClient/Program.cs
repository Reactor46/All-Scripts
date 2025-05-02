using Inventory.Model;
using System;
using Newtonsoft.Json;
using System.Net.Http;
using System.Text;
using System.Collections.Generic;
using System.Linq;

namespace Inventory.Client
{
    class Program
    {
        static void Main(string[] args)
        {
            var dbInventory = new Databases.Inventory();
            var vdirInventory = new VirtualDirectories.Inventory();
            var schTskInventory = new ScheduledTasks.Inventory();

            var s = new Server() {
                Id = Environment.MachineName,
                Timestamp = DateTime.Now
            };

            Console.WriteLine("Collecting databases ...");
            s.Databases = dbInventory.GetDatabases();
            Console.WriteLine("Collecting database jobs ...");
            s.DatabaseJobs = dbInventory.GetDatabaseJobs();
            Console.WriteLine("Collecting virtual directories ...");
            s.VirtualDirectories = vdirInventory.GetVirtualDirectories();
            Console.WriteLine("Collecting application pools ...");
            s.ApplicationPools = vdirInventory.GetApplicationPools();
            Console.WriteLine("Collecting scheduled tasks ...");
            s.ScheduledTasks = schTskInventory.GetScheduledTasks();


            var msg = string.Empty;

            var client = new HttpClient();

            var serverSpecificUrl = 
                Properties.Settings.Default.InventoryServerURL.TrimEnd("/".ToCharArray())
                + "/api/servers/" 
                + Environment.MachineName;

            Console.WriteLine($"Deleting current information about {Environment.MachineName}, if any ...");
            var x =
                client
                .DeleteAsync(serverSpecificUrl)
                .ContinueWith((response) => { msg = response.Result.Content.ReadAsStringAsync().Result; });

            x.Wait();
            Console.WriteLine($"Posting updated information about {Environment.MachineName} ...");
            x = client.PostAsync(serverSpecificUrl,
               new StringContent(JsonConvert.SerializeObject(s),
               Encoding.UTF8,
               "application/json"))
               .ContinueWith((response) => { msg = response.Result.Content.ReadAsStringAsync().Result; });

            x.Wait();

            Console.WriteLine(msg);

            Console.WriteLine("Done");
        }
    }
}
