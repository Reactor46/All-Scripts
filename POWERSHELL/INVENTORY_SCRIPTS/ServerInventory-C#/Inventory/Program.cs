using Microsoft.Owin.Hosting;
using System;

namespace Inventory.Server
{
    public class Program
    {

        static void Main(string[] args)
        {
            string baseUri = Properties.Settings.Default.InventoryServerURL;


            var context = new InventoryContext();
            Console.WriteLine(context.Database.Connection.ConnectionString);

            Console.WriteLine("Starting web Server...");
            WebApp.Start<Startup>(baseUri);
            Console.WriteLine($"Server running at {Environment.MachineName}, {baseUri} - press Enter to quit. ");
            Console.ReadLine();
        }
    }
}