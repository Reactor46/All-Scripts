﻿using System;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;
using System.Web;
using System.Web.SessionState;
using Powershell = System.Management.Automation.PowerShell;

namespace System.Web.Handlers
{
    class PowershellHandler : IHttpHandler, IRequiresSessionState
    {
        public bool IsReusable
        {
            get { return false; }
        }

        public void ProcessRequest(HttpContext context)
        {
            Runspace rs = this.getRunspace(context);
            using (Powershell ps = Powershell.Create())
            {
                ps.Runspace = rs;
                ps.AddScript(this.getFileContent(context.Request.PhysicalPath));
                try
                {
                    foreach (PSObject result in ps.Invoke())
                    {
                        context.Response.Write(result);
                    }

                    PSDataCollection<ErrorRecord> errors = ps.Streams.Error;
                    if (errors != null && errors.Count > 0)
                    {
                        foreach (ErrorRecord err in errors)
                        {
                            context.Response.Write(String.Format("  error: {0}<br />", err.ToString()));
                        }
                    }
                }
                catch (RuntimeException runtimeException)
                {
                    // Trap any exception generated by the commands. These exceptions
                    // will all be derived from the RuntimeException exception.
                    context.Response.Write(String.Format("Runtime exception: {0}: {1}<br />{2}<br />",
                        runtimeException.ErrorRecord.InvocationInfo.InvocationName,
                        runtimeException.Message,
                        runtimeException.ErrorRecord.InvocationInfo.PositionMessage));
                }
                ps.Dispose();
            }
        }

        private Runspace getRunspace(HttpContext context)
        {
            Runspace rs;

            if (context.Session["Runspace"] == null)
            {
                InitialSessionState iss = InitialSessionState.CreateDefault();
                iss.AuthorizationManager = new AuthorizationManager("PowershellHandler");
                rs = RunspaceFactory.CreateRunspace(iss);
                context.Session["Runspace"] = rs;
                rs.Open();
                this.configApp(context);
            }

            rs = context.Session["Runspace"] as Runspace;
            rs.SessionStateProxy.SetVariable("HttpContext", context);

            return rs;
        }

        private void configApp(HttpContext context)
        {
            String confPath = String.Format("{0}\\App_Code\\config.ps1", context.Request.PhysicalApplicationPath);
            if (File.Exists(confPath))
            {
                using (Powershell ps = Powershell.Create())
                {
                    ps.Runspace = context.Session["Runspace"] as Runspace;
                    ps.Runspace.SessionStateProxy.SetVariable("HttpContext", context);
                    ps.AddScript(this.getFileContent(confPath));
                    ps.Invoke();
                }
            }
        }

        private String getFileContent(String path)
        {
            String script = String.Empty;
            using (StreamReader streamReader = new StreamReader(path, Encoding.UTF8))
            {
                script = streamReader.ReadToEnd();
            }

            return script;
        }
    }
}