using Inventory.Model;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Management;
using System.Threading.Tasks;

namespace Inventory.Client.Databases
{
    public class Inventory
    {
        public List<Database> GetDatabases()
        {
            var result = new List<Database>();


            foreach (var connectionString in ConnectionStrings)
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    var cmd = new SqlCommand("select name, state_desc from sys.databases where len(owner_sid) > 1", conn);
                    var reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        result.Add(
                            new Database()
                            {
                                Name = reader.GetString(0),
                                Online = reader.GetString(1).Equals("ONLINE"),
                                Timestamp = DateTime.Now
                            });

                    }

                    conn.Close();
                }
            }

            return result;
        }

        public List<DatabaseJob> GetDatabaseJobs()
        {
            var result = new List<DatabaseJob>();

            foreach (var connectionString in ConnectionStrings)
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    var cmd = new SqlCommand("select name, enabled from [msdb].[dbo].[sysjobs]", conn);
                    var reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        result.Add(
                            new DatabaseJob()
                            {
                                Name = reader.GetString(0),
                                Enabled = (reader.GetByte(1) == 1),
                                Timestamp = DateTime.Now
                            });

                    }

                    conn.Close();
                }
            }

            return result;
        }

        private List<string> _ConnectionStrings;
        public List<string> ConnectionStrings
        {
            get
            {
                if (_ConnectionStrings == null)
                {
                    _ConnectionStrings = new List<string>();

                    var candidates = new List<string>();
                    candidates.Add($@"Data Source={Environment.MachineName};Integrated Security=True;Connection Timeout=10");



                    foreach (var sqlInstanceName in GetSqlInstanceNames())
                    {
                        candidates.Add($@"Data Source={ string.Join(@"\", new string[] { System.Environment.MachineName, sqlInstanceName })};Integrated Security=True;Connection Timeout=10");
                    }

                    foreach (var sqlClusterName in GetSqlClusterNames())
                    {
                        candidates.Add($@"Data Source={ sqlClusterName };Integrated Security=True;Connection Timeout=10");
                        foreach (var sqlInstanceName in GetSqlInstanceNames())
                        {
                            candidates.Add($@"Data Source={ string.Join(@"\", new string[] { sqlClusterName, sqlInstanceName })};Integrated Security=True;Connection Timeout=10");
                        }
                    }



                    Parallel.ForEach(candidates, connectionstring =>
                    {
                        try
                        {
                            using (var conn = new SqlConnection(connectionstring))
                            {
                                conn.Open();
                                conn.Close();
                                _ConnectionStrings.Add(connectionstring);
                                
                                Console.WriteLine($"   OK: {connectionstring}");
                                
                            }
                        }
                        catch (SqlException sqle)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"  NOK: {connectionstring}");
                            Console.ForegroundColor = ConsoleColor.White;
                        }

                    });
                }

                return _ConnectionStrings;
            }
        }

        public IEnumerable<string> GetSqlInstanceNames()
        {
            try
            {
                return Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL").GetValueNames();
            }
            catch (Exception)
            {
                return new List<string>();
            }
        }

        public IEnumerable<string> GetSqlClusterNames()
        {
            var result = new List<string>();

            try
            {
                var sqlKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Microsoft SQL Server");

                foreach (string keyName in sqlKey.GetSubKeyNames())
                {
                    var key = sqlKey.OpenSubKey(keyName);
                    foreach (string subKeyName in key.GetSubKeyNames())
                    {
                        if (subKeyName.Equals("Cluster"))
                        {
                            var subKey = key.OpenSubKey(subKeyName);
                            result.Add((string)subKey.GetValue("ClusterName"));
                        }
                    }
                }
            }
            catch (Exception e)
            {

            }

            return result;
        }
    }
}
