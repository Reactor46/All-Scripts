using Inventory.Client.IIS6;
using Inventory.Model;
using Microsoft.Web.Administration;
using System.Collections.Generic;
using System.DirectoryServices;
using System;

namespace Inventory.Client.VirtualDirectories
{
    public class Inventory
    {
        public List<VirtualDir> GetVirtualDirectories()
        {
            var result = new List<VirtualDir>();

            result.AddRange(GetIIS7VirtualDirectories());
            result.AddRange(GetIIS6VirtualDirectories());

            return result;
        }

        public List<AppPool> GetApplicationPools()
        {
            var result = new List<AppPool>();

            result.AddRange(GetIIS7ApplicationPools());
            result.AddRange(GetIIS6ApplicationPools());

            return result;
        }

        private IEnumerable<VirtualDir> GetIIS6VirtualDirectories()
        {
            var result = new List<VirtualDir>();

            try
            {
                var IIsEntities = new DirectoryEntry($"IIS://localhost/W3SVC");

                foreach (DirectoryEntry IIsEntity in IIsEntities.Children)
                {
                    if (IIsEntity.SchemaClassName == "IIsWebServer")
                    {
                        result.Add(
                            new VirtualDir()
                            {
                                Name = $"{IIsEntity.Name} - {IIsEntity.Properties["ServerComment"].Value.ToString()}",
                                Path = GetPath(IIsEntity),
                                State = ((ServerState)IIsEntity.Properties["ServerState"].Value).ToString(),
                                ApplicationPool = (string)IIsEntity.Properties["AppPoolId"].Value,
                                IISVersion = 6,
                                Timestamp = DateTime.Now
                            }
                            );
                    }
                }
            }
            catch (Exception e) { }

            return result;
        }

        private IEnumerable<AppPool> GetIIS6ApplicationPools()
        {
            var result = new List<AppPool>();

            try
            {

                var IIsAppPools = new DirectoryEntry(@"IIS://localhost/W3SVC/AppPools");


                foreach (DirectoryEntry IIsAppPool in IIsAppPools.Children)
                {
                    var intStatus = (int)IIsAppPool.InvokeGet("AppPoolState");
                    string status = string.Empty;
                    switch (intStatus)
                    {
                        case 2:
                            status = "Running";
                            break;
                        case 4:
                            status = "Stopped";
                            break;
                        default:
                            status = "Unknown";
                            break;
                    }

                    var idType = (int)IIsAppPool.InvokeGet("AppPoolIdentityType");
                    string user = string.Empty;

                    switch (idType)
                    {
                        case 0:
                            user = "Local System";
                            break;
                        case 1:
                            user = "Local Service";
                            break;
                        case 2:
                            user = "Network Service";
                            break;
                        case 3:
                            user = (string)IIsAppPool.InvokeGet("WAMUserName");
                            break;

                    }


                    result.Add(
                        new AppPool()
                        {
                            Name = IIsAppPool.Name,
                            State = status,
                            Username = user,
                            IISVersion = 6,
                            Timestamp = DateTime.Now
                        });

                }

            }
            catch (Exception e) { }

            return result;
        }

        private IEnumerable<VirtualDir> GetIIS7VirtualDirectories()
        {
            var result = new List<VirtualDir>();

            try
            {
                using (ServerManager sm = new ServerManager())
                {

                    foreach (var site in sm.Sites)
                    {
                        result.Add(
                            new VirtualDir()
                            {
                                Name = site.Name,
                                State = site.State.ToString(),
                                Path = site.Applications["/"].VirtualDirectories["/"].PhysicalPath,
                                ApplicationPool = site.ApplicationDefaults.ApplicationPoolName,
                                IISVersion = 7,
                                Timestamp = DateTime.Now
                            }
                            );
                    }

                }
            }
            catch (Exception e) { }

            return result;
        }

        private IEnumerable<AppPool> GetIIS7ApplicationPools()
        {
            var result = new List<AppPool>();

            try
            {
                using (ServerManager sm = new ServerManager())
                {

                    foreach (var appPool in sm.ApplicationPools)
                    {
                        result.Add(
                            new AppPool()
                            {
                                Name = appPool.Name,
                                State = appPool.State.ToString(),
                                Username = appPool.ProcessModel.UserName,
                                IISVersion = 7,
                                Timestamp = DateTime.Now

                            });
                    }
                }
            }
            catch (Exception e) { }

            return result;
        }

        private string GetPath(DirectoryEntry de)
        {
            foreach (DirectoryEntry IIsEntity in de.Children)
            {
                if (IIsEntity.SchemaClassName == "IIsWebVirtualDir")
                    return IIsEntity.Properties["Path"].Value.ToString();
            }
            return string.Empty;
        }
    }
}
