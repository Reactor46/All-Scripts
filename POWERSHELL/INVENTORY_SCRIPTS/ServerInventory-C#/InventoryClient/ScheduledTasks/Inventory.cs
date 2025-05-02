using Inventory.Model;
using Microsoft.Win32.TaskScheduler;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Inventory.Client.ScheduledTasks
{
    public class Inventory
    {
        public List<ScheduledTask> GetScheduledTasks()
        {
            var result = new List<ScheduledTask>();

            using (TaskService ts = new TaskService())
            {
                var tasks = ts.FindAllTasks(new Regex(".*"));
                foreach (var task in tasks)
                {
                    if (task.Definition.Principal.UserId != null &&
                        !task.Definition.Principal.UserId.Equals("SYSTEM") &&
                        !task.Definition.Principal.UserId.Equals("LOCAL SERVICE") &&
                        !task.Definition.Principal.UserId.Equals("NETWORK SERVICE" ))
                        {
                        result.Add(
                            new ScheduledTask
                            {
                                Name = task.Name,
                                Enabled = task.Enabled,
                                Path = string.Join(", ",task.Definition.Actions.Select(x=>x.ToString())),
                                Principal = task.Definition.Principal.UserId,
                                Timestamp = DateTime.Now

                            });
                    }
                    
                }
            }


            return result;
        }
    }
}
