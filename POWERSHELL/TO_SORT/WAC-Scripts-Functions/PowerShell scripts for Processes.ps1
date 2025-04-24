function Get-WACPVCimNamespaceWithinMicrosoftWindows {
<#

.SYNOPSIS
Gets Namespace information under root/Microsoft/Windows

.DESCRIPTION
Gets Namespace information under root/Microsoft/Windows

.ROLE
Readers

#>

##SkipCheck=true##

Param(
)

import-module CimCmdlets

Get-CimInstance -Namespace root/Microsoft/Windows -Query "SELECT * FROM __NAMESPACE"

}
## [END] Get-WACPVCimNamespaceWithinMicrosoftWindows ##
function Get-WACPVCimProcess {
<#

.SYNOPSIS
Gets Msft_MTProcess objects.

.DESCRIPTION
Gets Msft_MTProcess objects.

.ROLE
Readers

#>

##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/Microsoft/Windows/ManagementTools -ClassName Msft_MTProcess

}
## [END] Get-WACPVCimProcess ##
function Get-WACPVProcessDownlevel {
<#

.SYNOPSIS
Gets information about the processes running in downlevel computer.

.DESCRIPTION
Gets information about the processes running in downlevel computer.

.ROLE
Readers

#>
param
(
    [Parameter(Mandatory = $true)]
    [boolean]
    $isLocal
)

$NativeProcessInfo = @"

namespace SMT
{
    using Microsoft.Win32.SafeHandles;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Runtime.InteropServices;

    public class SystemProcess
    {
        public uint processId;
        public uint parentId;
        public string name;
        public string description;
        public string executablePath;
        public string userName;
        public string commandLine;
        public uint sessionId;
        public uint processStatus;
        public ulong cpuTime;
        public ulong cycleTime;
        public DateTime CreationDateTime;
        public ulong workingSetSize;
        public ulong peakWorkingSetSize;
        public ulong privateWorkingSetSize;
        public ulong sharedWorkingSetSize;
        public ulong commitCharge;
        public ulong pagedPool;
        public ulong nonPagedPool;
        public uint pageFaults;
        public uint basePriority;
        public uint handleCount;
        public uint threadCount;
        public uint userObjects;
        public uint gdiObjects;
        public ulong readOperationCount;
        public ulong writeOperationCount;
        public ulong otherOperationCount;
        public ulong readTransferCount;
        public ulong writeTransferCount;
        public ulong otherTransferCount;
        public bool elevated;
        public double cpuPercent;
        public uint operatingSystemContext;
        public uint platform;
        public double cyclePercent;
        public ushort uacVirtualization;
        public ushort dataExecutionPrevention;
        public bool isImmersive;
        public ushort intervalSeconds;
        public ushort deltaWorkingSetSize;
        public ushort deltaPageFaults;
        public bool hasChildWindow;
        public string processType;
        public string fileDescription;

        public SystemProcess(NativeMethods.SYSTEM_PROCESS_INFORMATION processInformation)
        {
            this.processId = (uint)processInformation.UniqueProcessId.ToInt32();
            this.name = Marshal.PtrToStringAuto(processInformation.ImageName.Buffer);
            this.cycleTime = processInformation.CycleTime;
            this.cpuTime = (ulong)(processInformation.KernelTime + processInformation.UserTime);
            this.sessionId = processInformation.SessionId;
            this.workingSetSize = (ulong)(processInformation.WorkingSetSize.ToInt64() / 1024);
            this.peakWorkingSetSize = (ulong)processInformation.PeakWorkingSetSize.ToInt64();
            this.privateWorkingSetSize = (ulong)processInformation.WorkingSetPrivateSize;
            this.sharedWorkingSetSize = (ulong)processInformation.WorkingSetSize.ToInt64() - this.privateWorkingSetSize;
            this.commitCharge = (ulong)processInformation.PrivatePageCount.ToInt64();
            this.pagedPool = (ulong)processInformation.QuotaPagedPoolUsage.ToInt64();
            this.nonPagedPool = (ulong)processInformation.QuotaNonPagedPoolUsage.ToInt64();
            this.pageFaults = processInformation.PageFaultCount;
            this.handleCount = processInformation.HandleCount;
            this.threadCount = processInformation.NumberOfThreads;
            this.readOperationCount = (ulong)processInformation.ReadOperationCount;
            this.writeOperationCount = (ulong)processInformation.WriteOperationCount;
            this.otherOperationCount = (ulong)processInformation.OtherOperationCount;
            this.readTransferCount = (ulong)processInformation.ReadTransferCount;
            this.writeTransferCount = (ulong)processInformation.WriteTransferCount;
            this.otherTransferCount = (ulong)processInformation.OtherTransferCount;
            this.processStatus = 0;

            if(processInformation.BasePriority <= 4)
            {
                this.basePriority = 0x00000040; //IDLE_PRIORITY_CLASS
            }
            else if (processInformation.BasePriority <= 6)
            {
                this.basePriority = 0x00004000; //BELOW_NORMAL_PRIORITY_CLASS
            }
            else if (processInformation.BasePriority <= 8)
            {
                this.basePriority = 0x00000020; //NORMAL_PRIORITY_CLASS
            }
            else if (processInformation.BasePriority <= 10)
            {
                this.basePriority = 0x00008000; //ABOVE_NORMAL_PRIORITY_CLASS
            }
            else if (processInformation.BasePriority <= 13)
            {
                this.basePriority = 0x00000080; //HIGH_PRIORITY_CLASS
            }
            else
            {
                this.basePriority = 0x00000100; //REALTIME_PRIORITY_CLASS
            }
        }
    }

    public static class NativeMethods
    {
        [StructLayout(LayoutKind.Sequential)]
        internal struct UNICODE_STRING
        {
            internal ushort Length;
            internal ushort MaximumLength;
            internal IntPtr Buffer;
        }

        [System.Runtime.InteropServices.StructLayout(LayoutKind.Sequential)]
        public struct SYSTEM_PROCESS_INFORMATION
        {
            internal uint NextEntryOffset;
            internal uint NumberOfThreads;
            internal long WorkingSetPrivateSize;
            internal uint HardFaultCount;
            internal uint NumberOfThreadsHighWatermark;
            internal ulong CycleTime;
            internal long CreateTime;
            internal long UserTime;
            internal long KernelTime;
            internal UNICODE_STRING ImageName;
            internal int BasePriority;
            internal IntPtr UniqueProcessId;
            internal IntPtr InheritedFromUniqueProcessId;
            internal uint HandleCount;
            internal uint SessionId;
            internal IntPtr UniqueProcessKey;
            internal IntPtr PeakVirtualSize;
            internal IntPtr VirtualSize;
            internal uint PageFaultCount;
            internal IntPtr PeakWorkingSetSize;
            internal IntPtr WorkingSetSize;
            internal IntPtr QuotaPeakPagedPoolUsage;
            internal IntPtr QuotaPagedPoolUsage;
            internal IntPtr QuotaPeakNonPagedPoolUsage;
            internal IntPtr QuotaNonPagedPoolUsage;
            internal IntPtr PagefileUsage;
            internal IntPtr PeakPagefileUsage;
            internal IntPtr PrivatePageCount;
            internal long ReadOperationCount;
            internal long WriteOperationCount;
            internal long OtherOperationCount;
            internal long ReadTransferCount;
            internal long WriteTransferCount;
            internal long OtherTransferCount;
        }

        public enum TOKEN_INFORMATION_CLASS
        {
            TokenElevation = 20,
            TokenVirtualizationAllowed = 23,
            TokenVirtualizationEnabled = 24
        }

        [Flags]
        public enum ProcessAccessFlags : uint
        {
            QueryInformation = 0x00000400,
            QueryLimitedInformation = 0x00001000,
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct TOKEN_ELEVATION
        {
            public Int32 TokenIsElevated;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct UAC_ALLOWED
        {
            public Int32 UacAllowed;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct UAC_ENABLED
        {
            public Int32 UacEnabled;
        }

        [DllImport("ntdll.dll")]
        internal static extern int NtQuerySystemInformation(int SystemInformationClass, IntPtr SystemInformation, int SystemInformationLength, out int ReturnLength);

        [DllImport("kernel32.dll")]
        public static extern IntPtr OpenProcess(ProcessAccessFlags DesiredAccess, [MarshalAs(UnmanagedType.Bool)] bool InheritHandle, int ProcessId);

        [System.Runtime.InteropServices.DllImport("advapi32", CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true)]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool OpenProcessToken(IntPtr hProcess, UInt32 desiredAccess, out Microsoft.Win32.SafeHandles.SafeWaitHandle hToken);

        [System.Runtime.InteropServices.DllImport("advapi32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true)]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool GetTokenInformation(SafeWaitHandle hToken, TOKEN_INFORMATION_CLASS tokenInfoClass, IntPtr pTokenInfo, Int32 tokenInfoLength, out Int32 returnLength);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern uint GetGuiResources(IntPtr hProcess, uint uiFlags);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CloseHandle(IntPtr hObject);

        internal const int SystemProcessInformation = 5;

        internal const int STATUS_INFO_LENGTH_MISMATCH = unchecked((int)0xC0000004);

        internal const uint TOKEN_QUERY = 0x0008;
    }

    public static class Process
    {
        public static IEnumerable<SystemProcess> Enumerate()
        {
            List<SystemProcess> process = new List<SystemProcess>();

            int bufferSize = 1024;

            IntPtr buffer = Marshal.AllocHGlobal(bufferSize);

            QuerySystemProcessInformation(ref buffer, ref bufferSize);

            long totalOffset = 0;

            while (true)
            {
                IntPtr currentPtr = (IntPtr)((long)buffer + totalOffset);

                NativeMethods.SYSTEM_PROCESS_INFORMATION pi = new NativeMethods.SYSTEM_PROCESS_INFORMATION();

                pi = (NativeMethods.SYSTEM_PROCESS_INFORMATION)Marshal.PtrToStructure(currentPtr, typeof(NativeMethods.SYSTEM_PROCESS_INFORMATION));

                process.Add(new SystemProcess(pi));

                if (pi.NextEntryOffset == 0)
                {
                    break;
                }

                totalOffset += pi.NextEntryOffset;
            }

            Marshal.FreeHGlobal(buffer);

            GetExtendedProcessInfo(process);

            return process;
        }

        private static void GetExtendedProcessInfo(List<SystemProcess> processes)
        {
            foreach(var process in processes)
            {
                IntPtr hProcess = GetProcessHandle(process);

                if(hProcess != IntPtr.Zero)
                {
                    try
                    {
                        process.elevated = IsElevated(hProcess);
                        process.userObjects = GetCountUserResources(hProcess);
                        process.gdiObjects = GetCountGdiResources(hProcess);
                        process.uacVirtualization = GetVirtualizationStatus(hProcess);
                    }
                    finally
                    {
                        NativeMethods.CloseHandle(hProcess);
                    }
                }
            }
        }

        private static uint GetCountGdiResources(IntPtr hProcess)
        {
            return NativeMethods.GetGuiResources(hProcess, 0);
        }
        private static uint GetCountUserResources(IntPtr hProcess)
        {
            return NativeMethods.GetGuiResources(hProcess, 1);
        }

        private static ushort GetVirtualizationStatus(IntPtr hProcess)
        {
            /* Virtualization status:
             * 0: Unknown
             * 1: Disabled
             * 2: Enabled
             * 3: Not Allowed
             */
            ushort virtualizationStatus = 0;

            try
            {
                if(!IsVirtualizationAllowed(hProcess))
                {
                    virtualizationStatus = 3;
                }
                else
                {
                    if(IsVirtualizationEnabled(hProcess))
                    {
                        virtualizationStatus = 2;
                    }
                    else
                    {
                        virtualizationStatus = 1;
                    }
                }
            }
            catch(Win32Exception)
            {
            }

            return virtualizationStatus;
        }

        private static bool IsVirtualizationAllowed(IntPtr hProcess)
        {
            bool uacVirtualizationAllowed = false;

            Microsoft.Win32.SafeHandles.SafeWaitHandle hToken = null;
            int cbUacAlowed = 0;
            IntPtr pUacAllowed = IntPtr.Zero;

            try
            {
                if (!NativeMethods.OpenProcessToken(hProcess, NativeMethods.TOKEN_QUERY, out hToken))
                {
                    throw new Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error());
                }

                cbUacAlowed = System.Runtime.InteropServices.Marshal.SizeOf(typeof(NativeMethods.UAC_ALLOWED));
                pUacAllowed = System.Runtime.InteropServices.Marshal.AllocHGlobal(cbUacAlowed);

                if (pUacAllowed == IntPtr.Zero)
                {
                    throw new Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error());
                }

                if (!NativeMethods.GetTokenInformation(hToken, NativeMethods.TOKEN_INFORMATION_CLASS.TokenVirtualizationAllowed, pUacAllowed, cbUacAlowed, out cbUacAlowed))
                {
                    throw new Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error());
                }

                NativeMethods.UAC_ALLOWED uacAllowed = (NativeMethods.UAC_ALLOWED)System.Runtime.InteropServices.Marshal.PtrToStructure(pUacAllowed, typeof(NativeMethods.UAC_ALLOWED));

                uacVirtualizationAllowed = (uacAllowed.UacAllowed != 0);
            }
            finally
            {
                if (hToken != null)
                {
                    hToken.Close();
                    hToken = null;
                }

                if (pUacAllowed != IntPtr.Zero)
                {
                    System.Runtime.InteropServices.Marshal.FreeHGlobal(pUacAllowed);
                    pUacAllowed = IntPtr.Zero;
                    cbUacAlowed = 0;
                }
            }

            return uacVirtualizationAllowed;
        }

        public static bool IsVirtualizationEnabled(IntPtr hProcess)
        {
            bool uacVirtualizationEnabled = false;

            Microsoft.Win32.SafeHandles.SafeWaitHandle hToken = null;
            int cbUacEnabled = 0;
            IntPtr pUacEnabled = IntPtr.Zero;

            try
            {
                if (!NativeMethods.OpenProcessToken(hProcess, NativeMethods.TOKEN_QUERY, out hToken))
                {
                    throw new Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error());
                }

                cbUacEnabled = System.Runtime.InteropServices.Marshal.SizeOf(typeof(NativeMethods.UAC_ENABLED));
                pUacEnabled = System.Runtime.InteropServices.Marshal.AllocHGlobal(cbUacEnabled);

                if (pUacEnabled == IntPtr.Zero)
                {
                    throw new Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error());
                }

                if (!NativeMethods.GetTokenInformation(hToken, NativeMethods.TOKEN_INFORMATION_CLASS.TokenVirtualizationEnabled, pUacEnabled, cbUacEnabled, out cbUacEnabled))
                {
                    throw new Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error());
                }

                NativeMethods.UAC_ENABLED uacEnabled = (NativeMethods.UAC_ENABLED)System.Runtime.InteropServices.Marshal.PtrToStructure(pUacEnabled, typeof(NativeMethods.UAC_ENABLED));

                uacVirtualizationEnabled = (uacEnabled.UacEnabled != 0);
            }
            finally
            {
                if (hToken != null)
                {
                    hToken.Close();
                    hToken = null;
                }

                if (pUacEnabled != IntPtr.Zero)
                {
                    System.Runtime.InteropServices.Marshal.FreeHGlobal(pUacEnabled);
                    pUacEnabled = IntPtr.Zero;
                    cbUacEnabled = 0;
                }
            }

            return uacVirtualizationEnabled;
        }

        private static bool IsElevated(IntPtr hProcess)
        {
             bool fIsElevated = false;
            Microsoft.Win32.SafeHandles.SafeWaitHandle hToken = null;
            int cbTokenElevation = 0;
            IntPtr pTokenElevation = IntPtr.Zero;

            try
            {
                if (!NativeMethods.OpenProcessToken(hProcess, NativeMethods.TOKEN_QUERY, out hToken))
                {
                    throw new Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error());
                }

                cbTokenElevation = System.Runtime.InteropServices.Marshal.SizeOf(typeof(NativeMethods.TOKEN_ELEVATION));
                pTokenElevation = System.Runtime.InteropServices.Marshal.AllocHGlobal(cbTokenElevation);

                if (pTokenElevation == IntPtr.Zero)
                {
                    throw new Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error());
                }

                if (!NativeMethods.GetTokenInformation(hToken, NativeMethods.TOKEN_INFORMATION_CLASS.TokenElevation, pTokenElevation, cbTokenElevation, out cbTokenElevation))
                {
                    throw new Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error());
                }

                NativeMethods.TOKEN_ELEVATION elevation = (NativeMethods.TOKEN_ELEVATION)System.Runtime.InteropServices.Marshal.PtrToStructure(pTokenElevation, typeof(NativeMethods.TOKEN_ELEVATION));

                fIsElevated = (elevation.TokenIsElevated != 0);
            }
            catch (Win32Exception)
            {
            }
            finally
            {
                if (hToken != null)
                {
                    hToken.Close();
                    hToken = null;
                }

                if (pTokenElevation != IntPtr.Zero)
                {
                    System.Runtime.InteropServices.Marshal.FreeHGlobal(pTokenElevation);
                    pTokenElevation = IntPtr.Zero;
                    cbTokenElevation = 0;
                }
            }

            return fIsElevated;
        }

        private static IntPtr GetProcessHandle(SystemProcess process)
        {
            IntPtr hProcess = NativeMethods.OpenProcess(NativeMethods.ProcessAccessFlags.QueryInformation | NativeMethods.ProcessAccessFlags.QueryLimitedInformation, false, (int)process.processId);

            if(hProcess == IntPtr.Zero)
            {
                hProcess = NativeMethods.OpenProcess(NativeMethods.ProcessAccessFlags.QueryLimitedInformation, false, (int)process.processId);
            }

            return hProcess;
        }

        private static void QuerySystemProcessInformation(ref IntPtr processInformationBuffer, ref int processInformationBufferSize)
        {
            const int maxTries = 10;
            bool success = false;

            for (int i = 0; i < maxTries; i++)
            {
                int sizeNeeded;

                int result = NativeMethods.NtQuerySystemInformation(NativeMethods.SystemProcessInformation, processInformationBuffer, processInformationBufferSize, out sizeNeeded);

                if (result == NativeMethods.STATUS_INFO_LENGTH_MISMATCH)
                {
                    if (processInformationBuffer != IntPtr.Zero)
                    {
                        Marshal.FreeHGlobal(processInformationBuffer);
                    }

                    processInformationBuffer = Marshal.AllocHGlobal(sizeNeeded);
                    processInformationBufferSize = sizeNeeded;
                }

                else if (result < 0)
                {
                    throw new Exception(String.Format("NtQuerySystemInformation failed with code 0x{0:X8}", result));
                }

                else
                {
                    success = true;
                    break;
                }
            }

            if (!success)
            {
                throw new Exception("Failed to allocate enough memory for NtQuerySystemInformation");
            }
        }
    }
}
"@

############################################################################################################################

# Global settings for the script.

############################################################################################################################

$ErrorActionPreference = "Stop"

Set-StrictMode -Version 3.0

############################################################################################################################

# Helper functions.

############################################################################################################################

function Get-ProcessListFromWmi {
    <#
    .Synopsis
        Name: Get-ProcessListFromWmi
        Description: Runs the WMI command to get Win32_Process objects and returns them in hashtable where key is processId.

    .Returns
        The list of processes in the form of hashtable.
    #>
    $processList = @{}

    $WmiProcessList = Get-WmiObject -Class Win32_Process

    foreach ($process in $WmiProcessList) {
        $processList.Add([int]$process.ProcessId, $process)
    }

    $processList
}

function Get-ProcessPerfListFromWmi {
    <#
    .Synopsis
        Name: Get-ProcessPerfListFromWmi
        Description: Runs the WMI command to get Win32_PerfFormattedData_PerfProc_Process objects and returns them in hashtable where key is processId.

    .Returns
        The list of processes performance data in the form of hashtable.
    #>
    $processPerfList = @{}

    $WmiProcessPerfList = Get-WmiObject -Class Win32_PerfFormattedData_PerfProc_Process

    foreach ($process in $WmiProcessPerfList) {
        try {
            $processPerfList.Add([int]$process.IdProcess, $process)
        }
        catch {
            if ($_.FullyQualifiedErrorId -eq 'ArgumentException') {
                $processPerfList.Remove([int]$process.IdProcess)
            }

            $processPerfList.Add([int]$process.IdProcess, $process)
        }
    }

    $processPerfList
}

function Get-ProcessListFromPowerShell {
    <#
    .Synopsis
        Name: Get-ProcessListFromPowerShell
        Description: Runs the PowerShell command Get-Process to get process objects.

    .Returns
        The list of processes in the form of hashtable.
    #>
    $processList = @{}

    if ($psVersionTable.psversion.Major -ge 4) {
        #
        # It will crash to run 'Get-Process' with parameter 'IncludeUserName' multiple times in a session.
        # Currently the UI will not reuse the session as a workaround.
        # We need to remove the paramter 'IncludeUserName' if this issue happens again.
        #
        $PowerShellProcessList = Get-Process -IncludeUserName -ErrorAction SilentlyContinue
    }
    else {
        $PowerShellProcessList = Get-Process -ErrorAction SilentlyContinue
    }

    foreach ($process in $PowerShellProcessList) {
        $processList.Add([int]$process.Id, $process)
    }

    $processList
}

function Get-LocalSystemAccount {
    <#
    .Synopsis
        Name: Get-LocalSystemAccount
        Description: Gets the name of local system account.

    .Returns
        The name local system account.
    #>
    $sidLocalSystemAccount = "S-1-5-18"

    $objSID = New-Object System.Security.Principal.SecurityIdentifier($sidLocalSystemAccount)

    $objSID.Translate( [System.Security.Principal.NTAccount]).Value
}

function Get-NumberOfLogicalProcessors {
    <#
    .Synopsis
        Name: Get-NumberOfLogicalProcessors
        Description: Gets the number of logical processors on the system.

    .Returns
        The number of logical processors on the system.
    #>
    $computerSystem = Get-CimInstance -Class Win32_ComputerSystem -Property NumberOfLogicalProcessors -ErrorAction Stop
    if ($computerSystem) {
        $computerSystem.NumberOfLogicalProcessors
    }
    else {
        throw 'Unable to get processor information'
    }
}


############################################################################################################################
# Main script.
############################################################################################################################

Add-Type -TypeDefinition $NativeProcessInfo
Remove-Variable NativeProcessInfo

try {
    #
    # Get the information about system processes from different sources.
    #
    $NumberOfLogicalProcessors = Get-NumberOfLogicalProcessors
    $NativeProcesses = [SMT.Process]::Enumerate()
    $WmiProcesses = Get-ProcessListFromWmi
    $WmiPerfProcesses = Get-ProcessPerfListFromWmi
    $PowerShellProcesses = Get-ProcessListFromPowerShell
    $LocalSystemAccount = Get-LocalSystemAccount

    $systemIdleProcess = $null
    $cpuInUse = 0

    # process paths and categorization taken from Task Manager
    # https://microsoft.visualstudio.com/_git/os?path=%2Fbase%2Fdiagnosis%2Fpdui%2Fatm%2FApplications.cpp&version=GBofficial%2Frs_fun_flight&_a=contents&line=44&lineStyle=plain&lineEnd=59&lineStartColumn=1&lineEndColumn=3
    $criticalProcesses = (
        "$($env:windir)\system32\winlogon.exe",
        "$($env:windir)\system32\wininit.exe",
        "$($env:windir)\system32\csrss.exe",
        "$($env:windir)\system32\lsass.exe",
        "$($env:windir)\system32\smss.exe",
        "$($env:windir)\system32\services.exe",
        "$($env:windir)\system32\taskeng.exe",
        "$($env:windir)\system32\taskhost.exe",
        "$($env:windir)\system32\dwm.exe",
        "$($env:windir)\system32\conhost.exe",
        "$($env:windir)\system32\svchost.exe",
        "$($env:windir)\system32\sihost.exe",
        "$($env:ProgramFiles)\Windows Defender\msmpeng.exe",
        "$($env:ProgramFiles)\Windows Defender\nissrv.exe",
        "$($env:windir)\explorer.exe"
    )

    $sidebarPath = "$($end:ProgramFiles)\Windows Sidebar\sidebar.exe"
    $appFrameHostPath = "$($env:windir)\system32\ApplicationFrameHost.exe"

    $edgeProcesses = (
        "$($env:windir)\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdge.exe",
        "$($env:windir)\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdgeCP.exe",
        "$($env:windir)\system32\browser_broker.exe"
    )

    #
    # Extract the additional process related information and fill up each nativeProcess object.
    #
    foreach ($nativeProcess in $NativeProcesses) {
        $WmiProcess = $null
        $WmiPerfProcess = $null
        $psProcess = $null

        # Same process as retrieved from WMI call Win32_Process
        if ($WmiProcesses.ContainsKey([int]$nativeProcess.ProcessId)) {
            $WmiProcess = $WmiProcesses.Get_Item([int]$nativeProcess.ProcessId)
        }

        # Same process as retrieved from WMI call Win32_PerfFormattedData_PerfProc_Process
        if ($WmiPerfProcesses.ContainsKey([int]$nativeProcess.ProcessId)) {
            $WmiPerfProcess = $WmiPerfProcesses.Get_Item([int]$nativeProcess.ProcessId)
        }

        # Same process as retrieved from PowerShell call Win32_Process
        if ($PowerShellProcesses.ContainsKey([int]$nativeProcess.ProcessId)) {
            $psProcess = $PowerShellProcesses.Get_Item([int]$nativeProcess.ProcessId)
        }

        if (($WmiProcess -eq $null) -or ($WmiPerfProcess -eq $null) -or ($psProcess -eq $null)) {continue}

        $nativeProcess.name = $WmiProcess.Name
        $nativeProcess.description = $WmiProcess.Description
        $nativeProcess.executablePath = $WmiProcess.ExecutablePath
        $nativeProcess.commandLine = $WmiProcess.CommandLine
        $nativeProcess.parentId = $WmiProcess.ParentProcessId

        #
        # Process CPU utilization and divide by number of cores
        # Win32_PerfFormattedData_PerfProc_Process PercentProcessorTime has a max number of 100 * cores so we want to normalize it
        #
        if ($WmiPerfProcess -and $WmiPerfProcess.PercentProcessorTime -ne $null -and $NumberOfLogicalProcessors -gt 0) {
            $nativeProcess.cpuPercent = $WmiPerfProcess.PercentProcessorTime / $NumberOfLogicalProcessors
        }
        #
        # Process start time.
        #
        if ($WmiProcess.CreationDate) {
            $nativeProcess.CreationDateTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($WmiProcess.CreationDate)
        }
        else {
            if ($nativeProcess.ProcessId -in @(0, 4)) {
                # Under some circumstances, the process creation time is not available for processs "System Idle Process" or "System"
                # In this case we assume that the process creation time is when the system was last booted.
                $nativeProcess.CreationDateTime = [System.Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject -Class win32_Operatingsystem).LastBootUpTime)
            }
        }

        #
        # Owner of the process.
        #
        if ($psVersionTable.psversion.Major -ge 4) {
            $nativeProcess.userName = $psProcess.UserName
        }

        # If UserName was not present available in results returned from Get-Process, then get the UserName from WMI class Get-Process
        <#
        ###### GetOwner is too slow so skip this part. ####

        if([string]::IsNullOrWhiteSpace($nativeProcess.userName))
        {
            $processOwner = Invoke-WmiMethod -InputObject $WmiProcess -Name GetOwner -ErrorAction SilentlyContinue

            try
            {
                if($processOwner.Domain)
                {
                    $nativeProcess.userName = "{0}\{1}" -f $processOwner.Domain, $processOwner.User
                }
                else
                {
                    $nativeProcess.userName = "{0}" -f $processOwner.User
                }
            }
            catch
            {
            }

            #In case of 'System Idle Process" and 'System' there is a need to explicitly mention NT Authority\System as Process Owner.
            if([string]::IsNullOrWhiteSpace($nativeProcess.userName) -and $nativeProcess.processId -in @(0, 4))
            {
                   $nativeProcess.userName = Get-LocalSystemAccount
            }
        }
        #>

        #In case of 'System Idle Process" and 'System' there is a need to explicitly mention NT Authority\System as Process Owner.
        if ([string]::IsNullOrWhiteSpace($nativeProcess.userName) -and $nativeProcess.processId -in @(0, 4)) {
            $nativeProcess.userName = $LocalSystemAccount
        }

        #
        # The process status ( i.e. running or suspended )
        #
        $countSuspendedThreads = @($psProcess.Threads | where { $_.WaitReason -eq [System.Diagnostics.ThreadWaitReason]::Suspended }).Count

        if ($psProcess.Threads.Count -eq $countSuspendedThreads) {
            $nativeProcess.ProcessStatus = 2
        }
        else {
            $nativeProcess.ProcessStatus = 1
        }

        # calculate system idle process
        if ($nativeProcess.processId -eq 0) {
            $systemIdleProcess = $nativeProcess
        }
        else {
            $cpuInUse += $nativeProcess.cpuPercent
        }


        if ($isLocal) {
            $nativeProcess.hasChildWindow = $psProcess -ne $null -and $psProcess.MainWindowHandle -ne 0

            if ($psProcess.MainModule -and $psProcess.MainModule.FileVersionInfo) {
                $nativeProcess.fileDescription = $psProcess.MainModule.FileVersionInfo.FileDescription
            }

            if ($edgeProcesses -contains $nativeProcess.executablePath) {
                # special handling for microsoft edge used by task manager
                # group all edge processes into applications
                $nativeProcess.fileDescription = 'Microsoft Edge'
                $nativeProcess.processType = 'application'
            }
            elseif ($criticalProcesses -contains $nativeProcess.executablePath `
                    -or (($nativeProcess.executablePath -eq $null -or $nativeProcess.executablePath -eq '') -and $null -ne ($criticalProcesses | ? {$_ -match $nativeProcess.name})) ) {
                # process is windows if its executable path is a critical process, defined by Task Manager
                # if the process has no executable path recorded, fallback to use the name to match to critical process
                $nativeProcess.processType = 'windows'
            }
            elseif (($nativeProcess.hasChildWindow -and $nativeProcess.executablePath -ne $appFrameHostPath) -or $nativeProcess.executablePath -eq $sidebarPath) {
                # sidebar.exe, or has child window (excluding ApplicationFrameHost.exe)
                $nativeProcess.processType = 'application'
            }
            else {
                $nativeProcess.processType = 'background'
            }
        }
    }

    if ($systemIdleProcess -ne $null) {
        $systemIdleProcess.cpuPercent = [Math]::Max(100 - $cpuInUse, 0)
    }

}
catch {
    throw $_
}
finally {
    $WmiProcesses = $null
    $WmiPerfProcesses = $null
}

# Return the result to the caller of this script.
$NativeProcesses


}
## [END] Get-WACPVProcessDownlevel ##
function Get-WACPVProcessHandle {
<#

.SYNOPSIS
Gets the filtered information of all the Operating System handles.

.DESCRIPTION
Gets the filtered information of all the Operating System handles.

.ROLE
Readers

#>

param (
    [Parameter(Mandatory = $true, ParameterSetName = 'processId')]
    [int]
    $processId,

    [Parameter(Mandatory = $true, ParameterSetName = 'handleSubstring')]
    [string]
    $handleSubstring
)

$SystemHandlesInfo = @"
    
namespace SME
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading;

    public static class NativeMethods
    {
        internal enum SYSTEM_INFORMATION_CLASS : int
        {
            /// </summary>
            SystemHandleInformation = 16
        }

        [Flags]
        internal enum ProcessAccessFlags : int
        {
            All = 0x001F0FFF,
            Terminate = 0x00000001,
            CreateThread = 0x00000002,
            VMOperation = 0x00000008,
            VMRead = 0x00000010,
            VMWrite = 0x00000020,
            DupHandle = 0x00000040,
            SetInformation = 0x00000200,
            QueryInformation = 0x00000400,
            QueryLimitedInformation = 0x00001000,
            Synchronize = 0x00100000
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct SystemHandle
        {
            public Int32 ProcessId;
            public Byte ObjectTypeNumber;
            public Byte Flags;
            public UInt16 Handle;
            public IntPtr Object;
            public Int32 GrantedAccess;
        }

        [Flags]
        public enum DuplicateOptions : int
        {
            NONE = 0,
            /// <summary>
            /// Closes the source handle. This occurs regardless of any error status returned.
            /// </summary>
            DUPLICATE_CLOSE_SOURCE = 0x00000001,
            /// <summary>
            /// Ignores the dwDesiredAccess parameter. The duplicate handle has the same access as the source handle.
            /// </summary>
            DUPLICATE_SAME_ACCESS = 0x00000002
        }

        internal enum OBJECT_INFORMATION_CLASS : int
        {
            /// <summary>
            /// Returns a PUBLIC_OBJECT_BASIC_INFORMATION structure as shown in the following Remarks section.
            /// </summary>
            ObjectBasicInformation = 0,
            ObjectNameInformation = 1,
            /// <summary>
            /// Returns a PUBLIC_OBJECT_TYPE_INFORMATION structure as shown in the following Remarks section.
            /// </summary>
            ObjectTypeInformation = 2
        }

        public enum FileType : int
        {
            FileTypeChar = 0x0002,
            FileTypeDisk = 0x0001,
            FileTypePipe = 0x0003,
            FileTypeRemote = 0x8000,
            FileTypeUnknown = 0x0000,
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct GENERIC_MAPPING
        {
            UInt32 GenericRead;
            UInt32 GenericWrite;
            UInt32 GenericExecute;
            UInt32 GenericAll;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct OBJECT_TYPE_INFORMATION
        {
            public UNICODE_STRING TypeName;
            public UInt32 TotalNumberOfObjects;
            public UInt32 TotalNumberOfHandles;
            public UInt32 TotalPagedPoolUsage;
            public UInt32 TotalNonPagedPoolUsage;
            public UInt32 TotalNamePoolUsage;
            public UInt32 TotalHandleTableUsage;
            public UInt32 HighWaterNumberOfObjects;
            public UInt32 HighWaterNumberOfHandles;
            public UInt32 HighWaterPagedPoolUsage;
            public UInt32 HighWaterNonPagedPoolUsage;
            public UInt32 HighWaterNamePoolUsage;
            public UInt32 HighWaterHandleTableUsage;
            public UInt32 InvalidAttributes;
            public GENERIC_MAPPING GenericMapping;
            public UInt32 ValidAccessMask;
            public Boolean SecurityRequired;
            public Boolean MaintainHandleCount;
            public UInt32 PoolType;
            public UInt32 DefaultPagedPoolCharge;
            public UInt32 DefaultNonPagedPoolCharge;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct UNICODE_STRING
        {
            public UInt16 Length;
            public UInt16 MaximumLength;
            [MarshalAs(UnmanagedType.LPWStr)]
            public String Buffer;
        }

        [DllImport("ntdll.dll")]
        internal static extern Int32 NtQuerySystemInformation(
            SYSTEM_INFORMATION_CLASS SystemInformationClass,
            IntPtr SystemInformation,
            Int32 SystemInformationLength,
            out Int32 ReturnedLength);

        [DllImport("kernel32.dll")]
        internal static extern IntPtr OpenProcess(
            ProcessAccessFlags dwDesiredAccess,
            [MarshalAs(UnmanagedType.Bool)] bool bInheritHandle,
            Int32 dwProcessId);

        [DllImport("ntdll.dll")]
        internal static extern UInt32 NtQueryObject(
            Int32 Handle,
            OBJECT_INFORMATION_CLASS ObjectInformationClass,
            IntPtr ObjectInformation,
            Int32 ObjectInformationLength,
            out Int32 ReturnLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool DuplicateHandle(
            IntPtr hSourceProcessHandle,
            IntPtr hSourceHandle,
            IntPtr hTargetProcessHandle,
            out IntPtr lpTargetHandle,
            UInt32 dwDesiredAccess,
            [MarshalAs(UnmanagedType.Bool)]
            bool bInheritHandle,
            DuplicateOptions dwOptions);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool QueryFullProcessImageName([In]IntPtr hProcess, [In]Int32 dwFlags, [Out]StringBuilder exeName, ref Int32 size);

        [DllImport("psapi.dll")]
        public static extern UInt32 GetModuleBaseName(IntPtr hProcess, IntPtr hModule, StringBuilder baseName, UInt32 size);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern UInt32 QueryDosDevice(String lpDeviceName, System.Text.StringBuilder lpTargetPath, Int32 ucchMax);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern FileType GetFileType(IntPtr hFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool CloseHandle(IntPtr hObject);

        internal const Int32 STATUS_INFO_LENGTH_MISMATCH = unchecked((Int32)0xC0000004L);
        internal const Int32 STATUS_SUCCESS = 0x00000000;
    }

    public class SystemHandles
    {
        private Queue<SystemHandle> systemHandles;
        private Int32 processId;
        String fileNameToMatch;
        Dictionary<Int32, IntPtr> processIdToHandle;
        Dictionary<Int32, String> processIdToImageName;
        private const Int32 GetObjectNameTimeoutMillis = 50;
        private Thread backgroundWorker;
        private static object syncRoot = new Object();

        public static IEnumerable<SystemHandle> EnumerateAllSystemHandles()
        {
            SystemHandles systemHandles = new SystemHandles();

            return systemHandles.Enumerate(HandlesEnumerationScope.AllSystemHandles);
        }
        public static IEnumerable<SystemHandle> EnumerateProcessSpecificHandles(Int32 processId)
        {
            SystemHandles systemHandles = new SystemHandles(processId);

            return systemHandles.Enumerate(HandlesEnumerationScope.ProcessSpecificHandles);
        }

        public static IEnumerable<SystemHandle> EnumerateMatchingFileNameHandles(String fileNameToMatch)
        {
            SystemHandles systemHandles = new SystemHandles(fileNameToMatch);

            return systemHandles.Enumerate(HandlesEnumerationScope.MatchingFileNameHandles);
        }

        private SystemHandles()
        { }

        public SystemHandles(Int32 processId)
        {
            this.processId = processId;
        }

        public SystemHandles(String fileNameToMatch)
        {
            this.fileNameToMatch = fileNameToMatch;
        }

        public IEnumerable<SystemHandle> Enumerate(HandlesEnumerationScope handlesEnumerationScope)
        {
            IEnumerable<SystemHandle> handles = null;

            this.backgroundWorker = new Thread(() => handles = Enumerate_Internal(handlesEnumerationScope));

            this.backgroundWorker.IsBackground = true;

            this.backgroundWorker.Start();

            return handles;
        }

        public bool IsBusy
        {
            get
            {
                return this.backgroundWorker.IsAlive;
            }
        }

        public bool WaitForEnumerationToComplete(int timeoutMillis)
        {
            return this.backgroundWorker.Join(timeoutMillis);
        }

        private IEnumerable<SystemHandle> Enumerate_Internal(HandlesEnumerationScope handlesEnumerationScope)
        {
            Int32 result;
            Int32 bufferLength = 1024;
            IntPtr buffer = Marshal.AllocHGlobal(bufferLength);
            Int32 requiredLength;
            Int64 handleCount;
            Int32 offset = 0;
            IntPtr currentHandlePtr = IntPtr.Zero;
            NativeMethods.SystemHandle systemHandleStruct;
            Int32 systemHandleStructSize = 0;
            this.systemHandles = new Queue<SystemHandle>();
            this.processIdToHandle = new Dictionary<Int32, IntPtr>();
            this.processIdToImageName = new Dictionary<Int32, String>();

            while (true)
            {
                result = NativeMethods.NtQuerySystemInformation(
                    NativeMethods.SYSTEM_INFORMATION_CLASS.SystemHandleInformation,
                    buffer,
                    bufferLength,
                    out requiredLength);

                if (result == NativeMethods.STATUS_SUCCESS)
                {
                    break;
                }
                else if (result == NativeMethods.STATUS_INFO_LENGTH_MISMATCH)
                {
                    Marshal.FreeHGlobal(buffer);
                    bufferLength *= 2;
                    buffer = Marshal.AllocHGlobal(bufferLength);
                }
                else
                {
                    throw new InvalidOperationException(
                        String.Format(CultureInfo.InvariantCulture, "NtQuerySystemInformation failed with error code {0}", result));
                }
            } // End while loop.

            if (IntPtr.Size == 4)
            {
                handleCount = Marshal.ReadInt32(buffer);
            }
            else
            {
                handleCount = Marshal.ReadInt64(buffer);
            }

            offset = IntPtr.Size;
            systemHandleStruct = new NativeMethods.SystemHandle();
            systemHandleStructSize = Marshal.SizeOf(systemHandleStruct);

            if (handlesEnumerationScope == HandlesEnumerationScope.AllSystemHandles)
            {
                EnumerateAllSystemHandles(buffer, offset, systemHandleStructSize, handleCount);
            }
            else if (handlesEnumerationScope == HandlesEnumerationScope.ProcessSpecificHandles)
            {
                EnumerateProcessSpecificSystemHandles(buffer, offset, systemHandleStructSize, handleCount);
            }
            else if (handlesEnumerationScope == HandlesEnumerationScope.MatchingFileNameHandles)
            {
                this.EnumerateMatchingFileNameHandles(buffer, offset, systemHandleStructSize, handleCount);
            }

            if (buffer != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(buffer);
            }

            this.Cleanup();

            return this.systemHandles;
        }

        public IEnumerable<SystemHandle> ExtractResults()
        {
            lock (syncRoot)
            {
                while (this.systemHandles.Count > 0)
                {
                    yield return this.systemHandles.Dequeue();
                }
            }
        }

        private void EnumerateAllSystemHandles(IntPtr buffer, Int32 offset, Int32 systemHandleStructSize, Int64 handleCount)
        {
            for (Int64 i = 0; i < handleCount; i++)
            {
                NativeMethods.SystemHandle currentHandleInfo =
                        (NativeMethods.SystemHandle)Marshal.PtrToStructure((IntPtr)((Int64)buffer + offset), typeof(NativeMethods.SystemHandle));

                ExamineCurrentHandle(currentHandleInfo);

                offset += systemHandleStructSize;
            }
        }

        private void EnumerateProcessSpecificSystemHandles(IntPtr buffer, Int32 offset, Int32 systemHandleStructSize, Int64 handleCount)
        {
            for (Int64 i = 0; i < handleCount; i++)
            {
                NativeMethods.SystemHandle currentHandleInfo =
                        (NativeMethods.SystemHandle)Marshal.PtrToStructure((IntPtr)((Int64)buffer + offset), typeof(NativeMethods.SystemHandle));

                if (currentHandleInfo.ProcessId == this.processId)
                {
                    ExamineCurrentHandle(currentHandleInfo);
                }

                offset += systemHandleStructSize;
            }
        }

        private void EnumerateMatchingFileNameHandles(IntPtr buffer, Int32 offset, Int32 systemHandleStructSize, Int64 handleCount)
        {
            for (Int64 i = 0; i < handleCount; i++)
            {
                NativeMethods.SystemHandle currentHandleInfo =
                        (NativeMethods.SystemHandle)Marshal.PtrToStructure((IntPtr)((Int64)buffer + offset), typeof(NativeMethods.SystemHandle));

                ExamineCurrentHandleForForMatchingFileName(currentHandleInfo, this.fileNameToMatch);

                offset += systemHandleStructSize;
            }
        }

        private void ExamineCurrentHandle(
            NativeMethods.SystemHandle currentHandleInfo)
        {
            IntPtr sourceProcessHandle = this.GetProcessHandle(currentHandleInfo.ProcessId);

            if (sourceProcessHandle == IntPtr.Zero)
            {
                return;
            }

            String processImageName = this.GetProcessImageName(currentHandleInfo.ProcessId, sourceProcessHandle);

            IntPtr duplicateHandle = CreateDuplicateHandle(sourceProcessHandle, (IntPtr)currentHandleInfo.Handle);

            if (duplicateHandle == IntPtr.Zero)
            {
                return;
            }

            String objectType = GetObjectType(duplicateHandle);

            String objectName = String.Empty;

            if (objectType != "File")
            {
                objectName = GetObjectName(duplicateHandle);
            }
            else
            {
                Thread getObjectNameThread = new Thread(() => objectName = GetObjectName(duplicateHandle));
                getObjectNameThread.IsBackground = true;
                getObjectNameThread.Start();

                if (false == getObjectNameThread.Join(GetObjectNameTimeoutMillis))
                {
                    getObjectNameThread.Abort();

                    getObjectNameThread.Join(GetObjectNameTimeoutMillis);

                    objectName = String.Empty;
                }
                else
                {
                    objectName = GetRegularFileName(objectName);
                }

                getObjectNameThread = null;
            }

            if (!String.IsNullOrWhiteSpace(objectType) &&
                !String.IsNullOrWhiteSpace(objectName))
            {
                SystemHandle systemHandle = new SystemHandle();
                systemHandle.TypeName = objectType;
                systemHandle.Name = objectName;
                systemHandle.ObjectTypeNumber = currentHandleInfo.ObjectTypeNumber;
                systemHandle.ProcessId = currentHandleInfo.ProcessId;
                systemHandle.ProcessImageName = processImageName;

                RegisterHandle(systemHandle);
            }

            NativeMethods.CloseHandle(duplicateHandle);
        }

        private void ExamineCurrentHandleForForMatchingFileName(
             NativeMethods.SystemHandle currentHandleInfo, String fileNameToMatch)
        {
            IntPtr sourceProcessHandle = this.GetProcessHandle(currentHandleInfo.ProcessId);

            if (sourceProcessHandle == IntPtr.Zero)
            {
                return;
            }

            String processImageName = this.GetProcessImageName(currentHandleInfo.ProcessId, sourceProcessHandle);

            if (String.IsNullOrWhiteSpace(processImageName))
            {
                return;
            }

            IntPtr duplicateHandle = CreateDuplicateHandle(sourceProcessHandle, (IntPtr)currentHandleInfo.Handle);

            if (duplicateHandle == IntPtr.Zero)
            {
                return;
            }

            String objectType = GetObjectType(duplicateHandle);

            String objectName = String.Empty;

            Thread getObjectNameThread = new Thread(() => objectName = GetObjectName(duplicateHandle));

            getObjectNameThread.IsBackground = true;

            getObjectNameThread.Start();

            if (false == getObjectNameThread.Join(GetObjectNameTimeoutMillis))
            {
                getObjectNameThread.Abort();

                getObjectNameThread.Join(GetObjectNameTimeoutMillis);

                objectName = String.Empty;
            }
            else
            {
                objectName = GetRegularFileName(objectName);
            }

            getObjectNameThread = null;


            if (!String.IsNullOrWhiteSpace(objectType) &&
                !String.IsNullOrWhiteSpace(objectName))
            {
                if (objectName.ToLower().Contains(fileNameToMatch.ToLower()))
                {
                    SystemHandle systemHandle = new SystemHandle();
                    systemHandle.TypeName = objectType;
                    systemHandle.Name = objectName;
                    systemHandle.ObjectTypeNumber = currentHandleInfo.ObjectTypeNumber;
                    systemHandle.ProcessId = currentHandleInfo.ProcessId;
                    systemHandle.ProcessImageName = processImageName;

                    RegisterHandle(systemHandle);
                }
            }

            NativeMethods.CloseHandle(duplicateHandle);
        }

        private void RegisterHandle(SystemHandle systemHandle)
        {
            lock (syncRoot)
            {
                this.systemHandles.Enqueue(systemHandle);
            }
        }

        private String GetObjectName(IntPtr duplicateHandle)
        {
            String objectName = String.Empty;
            IntPtr objectNameBuffer = IntPtr.Zero;

            try
            {
                Int32 objectNameBufferSize = 0x1000;
                objectNameBuffer = Marshal.AllocHGlobal(objectNameBufferSize);
                Int32 actualObjectNameLength;

                UInt32 queryObjectNameResult = NativeMethods.NtQueryObject(
                    duplicateHandle.ToInt32(),
                    NativeMethods.OBJECT_INFORMATION_CLASS.ObjectNameInformation,
                    objectNameBuffer,
                    objectNameBufferSize,
                    out actualObjectNameLength);

                if (queryObjectNameResult != 0 && actualObjectNameLength > 0)
                {
                    Marshal.FreeHGlobal(objectNameBuffer);
                    objectNameBufferSize = actualObjectNameLength;
                    objectNameBuffer = Marshal.AllocHGlobal(objectNameBufferSize);

                    queryObjectNameResult = NativeMethods.NtQueryObject(
                        duplicateHandle.ToInt32(),
                        NativeMethods.OBJECT_INFORMATION_CLASS.ObjectNameInformation,
                        objectNameBuffer,
                        objectNameBufferSize,
                        out actualObjectNameLength);
                }

                // Get the name
                if (queryObjectNameResult == 0)
                {
                    NativeMethods.UNICODE_STRING name = (NativeMethods.UNICODE_STRING)Marshal.PtrToStructure(objectNameBuffer, typeof(NativeMethods.UNICODE_STRING));

                    objectName = name.Buffer;
                }
            }
            catch (ThreadAbortException)
            {
            }
            finally
            {
                if (objectNameBuffer != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(objectNameBuffer);
                }
            }

            return objectName;
        }

        private String GetObjectType(IntPtr duplicateHandle)
        {
            String objectType = String.Empty;

            Int32 objectTypeBufferSize = 0x1000;
            IntPtr objectTypeBuffer = Marshal.AllocHGlobal(objectTypeBufferSize);
            Int32 actualObjectTypeLength;

            UInt32 queryObjectResult = NativeMethods.NtQueryObject(
                duplicateHandle.ToInt32(),
                NativeMethods.OBJECT_INFORMATION_CLASS.ObjectTypeInformation,
                objectTypeBuffer,
                objectTypeBufferSize,
                out actualObjectTypeLength);

            if (queryObjectResult == 0)
            {
                NativeMethods.OBJECT_TYPE_INFORMATION typeInfo = (NativeMethods.OBJECT_TYPE_INFORMATION)Marshal.PtrToStructure(objectTypeBuffer, typeof(NativeMethods.OBJECT_TYPE_INFORMATION));

                objectType = typeInfo.TypeName.Buffer;
            }

            if (objectTypeBuffer != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(objectTypeBuffer);
            }

            return objectType;
        }

        private IntPtr GetProcessHandle(Int32 processId)
        {
            if (this.processIdToHandle.ContainsKey(processId))
            {
                return this.processIdToHandle[processId];
            }

            IntPtr processHandle = NativeMethods.OpenProcess
                (NativeMethods.ProcessAccessFlags.DupHandle | NativeMethods.ProcessAccessFlags.QueryInformation | NativeMethods.ProcessAccessFlags.VMRead, false, processId);

            if (processHandle != IntPtr.Zero)
            {
                this.processIdToHandle.Add(processId, processHandle);
            }
            else
            {
                // throw new Win32Exception(Marshal.GetLastWin32Error());
                //  Console.WriteLine("UNABLE TO OPEN PROCESS {0}", processId);
            }

            return processHandle;
        }

        private String GetProcessImageName(Int32 processId, IntPtr handleToProcess)
        {
            if (this.processIdToImageName.ContainsKey(processId))
            {
                return this.processIdToImageName[processId];
            }

            Int32 bufferSize = 1024;

            String strProcessImageName = String.Empty;

            StringBuilder processImageName = new StringBuilder(bufferSize);

            NativeMethods.QueryFullProcessImageName(handleToProcess, 0, processImageName, ref bufferSize);

            strProcessImageName = processImageName.ToString();

            if (!String.IsNullOrWhiteSpace(strProcessImageName))
            {
                try
                {
                    strProcessImageName = Path.GetFileName(strProcessImageName);
                }
                catch
                {
                }

                this.processIdToImageName.Add(processId, strProcessImageName);
            }

            return strProcessImageName;
        }

        private IntPtr CreateDuplicateHandle(IntPtr sourceProcessHandle, IntPtr handleToDuplicate)
        {
            IntPtr currentProcessHandle = Process.GetCurrentProcess().Handle;

            IntPtr duplicateHandle = IntPtr.Zero;

            NativeMethods.DuplicateHandle(
                sourceProcessHandle,
                handleToDuplicate,
                currentProcessHandle,
                out duplicateHandle,
                0,
                false,
                NativeMethods.DuplicateOptions.DUPLICATE_SAME_ACCESS);

            return duplicateHandle;
        }

        private static String GetRegularFileName(String deviceFileName)
        {
            String actualFileName = String.Empty;

            if (!String.IsNullOrWhiteSpace(deviceFileName))
            {
                foreach (var logicalDrive in Environment.GetLogicalDrives())
                {
                    StringBuilder targetPath = new StringBuilder(4096);

                    if (0 == NativeMethods.QueryDosDevice(logicalDrive.Substring(0, 2), targetPath, 4096))
                    {
                        return targetPath.ToString();
                    }

                    String targetPathStr = targetPath.ToString();

                    if (deviceFileName.StartsWith(targetPathStr))
                    {
                        actualFileName = deviceFileName.Replace(targetPathStr, logicalDrive.Substring(0, 2));

                        break;

                    }
                }

                if (String.IsNullOrWhiteSpace(actualFileName))
                {
                    actualFileName = deviceFileName;
                }
            }

            return actualFileName;
        }

        private void Cleanup()
        {
            foreach (var processHandle in this.processIdToHandle.Values)
            {
                NativeMethods.CloseHandle(processHandle);
            }

            this.processIdToHandle.Clear();
        }
    }

    public class SystemHandle
    {
        public String Name { get; set; }
        public String TypeName { get; set; }
        public byte ObjectTypeNumber { get; set; }
        public Int32 ProcessId { get; set; }
        public String ProcessImageName { get; set; }
    }
  
    public enum HandlesEnumerationScope
    {
        AllSystemHandles,
        ProcessSpecificHandles,
        MatchingFileNameHandles
    }
}
"@

############################################################################################################################

# Global settings for the script.

############################################################################################################################

$ErrorActionPreference = "Stop"

Set-StrictMode -Version 3.0

############################################################################################################################

# Main script.

############################################################################################################################


Add-Type -TypeDefinition $SystemHandlesInfo

Remove-Variable SystemHandlesInfo

if ($PSCmdlet.ParameterSetName -eq 'processId' -and $processId -ne $null) {

       $systemHandlesFinder = New-Object -TypeName SME.SystemHandles -ArgumentList $processId

       $scope = [SME.HandlesEnumerationScope]::ProcessSpecificHandles
}

elseif ($PSCmdlet.ParameterSetName -eq 'handleSubString') {
    
       $SystemHandlesFinder = New-Object -TypeName SME.SystemHandles -ArgumentList $handleSubstring

       $scope = [SME.HandlesEnumerationScope]::MatchingFileNameHandles
}


$SystemHandlesFinder.Enumerate($scope) | out-null

while($SystemHandlesFinder.IsBusy)
{
    $SystemHandlesFinder.ExtractResults() | Write-Output
    $SystemHandlesFinder.WaitForEnumerationToComplete(50) | out-null
}

$SystemHandlesFinder.ExtractResults() | Write-Output
}
## [END] Get-WACPVProcessHandle ##
function Get-WACPVProcessModule {
<#

.SYNOPSIS
Gets services associated with the process.

.DESCRIPTION
Gets services associated with the process.

.ROLE
Readers

#>

param (
    [Parameter(Mandatory=$true)]
    [UInt32]
    $processId
)

$process = Get-Process -PID $processId
$process.Modules | Microsoft.PowerShell.Utility\Select-Object ModuleName, FileVersion, FileName, @{Name="Image"; Expression={$process.Name}}, @{Name="PID"; Expression={$process.id}}


}
## [END] Get-WACPVProcessModule ##
function Get-WACPVProcessService {
<#

.SYNOPSIS
Gets services associated with the process.

.DESCRIPTION
Gets services associated with the process.

.ROLE
Readers

#>

param (
    [Parameter(Mandatory=$true)]
    [Int32]
    $processId
)

Import-Module CimCmdlets -ErrorAction SilentlyContinue

Get-CimInstance -ClassName Win32_service | Where-Object {$_.ProcessId -eq $processId} | Microsoft.PowerShell.Utility\Select-Object Name, processId, Description, Status, StartName



}
## [END] Get-WACPVProcessService ##
function Get-WACPVProcesses {
<#

.SYNOPSIS
Gets information about the processes running in computer.

.DESCRIPTION
Gets information about the processes running in computer.

.ROLE
Readers

.COMPONENT
ProcessList_Body

#>
param
(
    [Parameter(Mandatory = $true)]
    [boolean]
    $isLocal
)

Import-Module CimCmdlets -ErrorAction SilentlyContinue

$processes = Get-CimInstance -Namespace root/Microsoft/Windows/ManagementTools -ClassName Msft_MTProcess

$powershellProcessList = @{}
$powerShellProcesses = Get-Process -ErrorAction SilentlyContinue

foreach ($process in $powerShellProcesses) {
    $powershellProcessList.Add([int]$process.Id, $process)
}

if ($isLocal) {
    # critical processes taken from task manager code
    # https://microsoft.visualstudio.com/_git/os?path=%2Fbase%2Fdiagnosis%2Fpdui%2Fatm%2FApplications.cpp&version=GBofficial%2Frs_fun_flight&_a=contents&line=44&lineStyle=plain&lineEnd=59&lineStartColumn=1&lineEndColumn=3
    $criticalProcesses = (
        "$($env:windir)\system32\winlogon.exe",
        "$($env:windir)\system32\wininit.exe",
        "$($env:windir)\system32\csrss.exe",
        "$($env:windir)\system32\lsass.exe",
        "$($env:windir)\system32\smss.exe",
        "$($env:windir)\system32\services.exe",
        "$($env:windir)\system32\taskeng.exe",
        "$($env:windir)\system32\taskhost.exe",
        "$($env:windir)\system32\dwm.exe",
        "$($env:windir)\system32\conhost.exe",
        "$($env:windir)\system32\svchost.exe",
        "$($env:windir)\system32\sihost.exe",
        "$($env:ProgramFiles)\Windows Defender\msmpeng.exe",
        "$($env:ProgramFiles)\Windows Defender\nissrv.exe",
        "$($env:ProgramFiles)\Windows Defender\nissrv.exe",
        "$($env:windir)\explorer.exe"
    )

    $sidebarPath = "$($end:ProgramFiles)\Windows Sidebar\sidebar.exe"
    $appFrameHostPath = "$($env:windir)\system32\ApplicationFrameHost.exe"

    $edgeProcesses = (
        "$($env:windir)\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdge.exe",
        "$($env:windir)\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdgeCP.exe",
        "$($env:windir)\system32\browser_broker.exe"
    )

    foreach ($process in $processes) {

        if ($powershellProcessList.ContainsKey([int]$process.ProcessId)) {
            $psProcess = $powershellProcessList.Get_Item([int]$process.ProcessId)
            $hasChildWindow = $psProcess -ne $null -and $psProcess.MainWindowHandle -ne 0
            $process | Add-Member -MemberType NoteProperty -Name "HasChildWindow" -Value $hasChildWindow
            if ($psProcess.MainModule -and $psProcess.MainModule.FileVersionInfo) {
                $process | Add-Member -MemberType NoteProperty -Name "FileDescription" -Value $psProcess.MainModule.FileVersionInfo.FileDescription
            }
        }

        if ($edgeProcesses -contains $nativeProcess.executablePath) {
            # special handling for microsoft edge used by task manager
            # group all edge processes into applications
            $edgeLabel = 'Microsoft Edge'
            if ($process.fileDescription) {
                $process.fileDescription = $edgeLabel
            }
            else {
                $process | Add-Member -MemberType NoteProperty -Name "FileDescription" -Value $edgeLabel
            }

            $processType = 'application'
        }
        elseif ($criticalProcesses -contains $nativeProcess.executablePath `
                -or (($nativeProcess.executablePath -eq $null -or $nativeProcess.executablePath -eq '') -and $null -ne ($criticalProcesses | ? {$_ -match $nativeProcess.name})) ) {
            # process is windows if its executable path is a critical process, defined by Task Manager
            # if the process has no executable path recorded, fallback to use the name to match to critical process
            $processType = 'windows'
        }
        elseif (($nativeProcess.hasChildWindow -and $nativeProcess.executablePath -ne $appFrameHostPath) -or $nativeProcess.executablePath -eq $sidebarPath) {
            # sidebar.exe, or has child window (excluding ApplicationFrameHost.exe)
            $processType = 'application'
        }
        else {
            $processType = 'background'
        }

        $process | Add-Member -MemberType NoteProperty -Name "ProcessType" -Value $processType
    }
}

$processes

}
## [END] Get-WACPVProcesses ##
function New-WACPVCimProcessDump {
<#

.SYNOPSIS
Creates a new process dump.

.DESCRIPTION
Creates a new process dump.

.ROLE
Administrators

#>

##SkipCheck=true##

Param(
[System.UInt16]$ProcessId
)

import-module CimCmdlets

$keyInstance = New-CimInstance -Namespace root/Microsoft/Windows/ManagementTools -ClassName MSFT_MTProcess -Key @('ProcessId') -Property @{ProcessId=$ProcessId;} -ClientOnly
Invoke-CimMethod $keyInstance -MethodName CreateDump

}
## [END] New-WACPVCimProcessDump ##
function New-WACPVProcessDumpDownlevel {
<#

.SYNOPSIS
Creates the mini dump of the process on downlevel computer.

.DESCRIPTION
Creates the mini dump of the process on downlevel computer.

.ROLE
Administrators

#>

param
(
    # The process ID of the process whose mini dump is supposed to be created.
    [int]
    $processId,

    # Path to the process dump file name.
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String]
    $fileName
)

$NativeCode = @"

    namespace SME
    {
        using System;
        using System.Diagnostics;
        using System.Runtime.InteropServices;

        public static class ProcessMiniDump
        {
            private enum MINIDUMP_TYPE
            {
                MiniDumpNormal = 0x00000000,
                MiniDumpWithDataSegs = 0x00000001,
                MiniDumpWithFullMemory = 0x00000002,
                MiniDumpWithHandleData = 0x00000004,
                MiniDumpFilterMemory = 0x00000008,
                MiniDumpScanMemory = 0x00000010,
                MiniDumpWithUnloadedModules = 0x00000020,
                MiniDumpWithIndirectlyReferencedMemory = 0x00000040,
                MiniDumpFilterModulePaths = 0x00000080,
                MiniDumpWithProcessThreadData = 0x00000100,
                MiniDumpWithPrivateReadWriteMemory = 0x00000200,
                MiniDumpWithoutOptionalData = 0x00000400,
                MiniDumpWithFullMemoryInfo = 0x00000800,
                MiniDumpWithThreadInfo = 0x00001000,
                MiniDumpWithCodeSegs = 0x00002000
            };

            [DllImport("dbghelp.dll", CallingConvention = CallingConvention.Winapi, SetLastError = true)]
            private extern static bool MiniDumpWriteDump(
                System.IntPtr hProcess,
                int processId,
                Microsoft.Win32.SafeHandles.SafeFileHandle hFile,
                MINIDUMP_TYPE dumpType,
                System.IntPtr exceptionParam,
                System.IntPtr userStreamParam,
                System.IntPtr callbackParam);

            public static void Create(int processId, string fileName)
            {
                if (string.IsNullOrWhiteSpace(fileName))
                {
                    throw new ArgumentNullException(fileName);
                }

                if (processId < 0)
                {
                    throw new ArgumentException("Incorrect value of ProcessId", "processId");
                }

                System.IO.FileStream fileStream = null;

                try
                {
                    fileStream = System.IO.File.OpenWrite(fileName);
                    var proc = Process.GetProcessById(processId);

                    bool success = MiniDumpWriteDump(
                        proc.Handle,
                        proc.Id,
                        fileStream.SafeFileHandle,
                        MINIDUMP_TYPE.MiniDumpWithFullMemory | MINIDUMP_TYPE.MiniDumpWithFullMemoryInfo | MINIDUMP_TYPE.MiniDumpWithHandleData | MINIDUMP_TYPE.MiniDumpWithUnloadedModules | MINIDUMP_TYPE.MiniDumpWithThreadInfo,
                        IntPtr.Zero,
                        IntPtr.Zero,
                        IntPtr.Zero);

                    if (!success)
                    {
                        Marshal.ThrowExceptionForHR(Marshal.GetHRForLastWin32Error());
                    }
                }
                finally
                {
                    if (fileStream != null)
                    {
                        fileStream.Close();
                    }
                }
            }
        }
}

"@

############################################################################################################################

# Global settings for the script.

############################################################################################################################

$ErrorActionPreference = "Stop"

Set-StrictMode -Version 3.0

############################################################################################################################

# Main script.

############################################################################################################################

Add-Type -TypeDefinition $NativeCode
Remove-Variable NativeCode

$fileName = "$($env:temp)\$($fileName)"

try {
    # Create the mini dump using native call.
    try {
        [SME.ProcessMiniDump]::Create($processId, $fileName)
        $result = New-Object PSObject
        $result | Add-Member -MemberType NoteProperty -Name 'DumpFilePath' -Value $fileName
        $result
    }
    catch {
        if ($_.FullyQualifiedErrorId -eq "ArgumentException") {
            throw "Unable to create the mini dump of the process. Please make sure that the processId is correct and the user has required permissions to create the mini dump of the process."
        }
        elseif ($_.FullyQualifiedErrorId -eq "UnauthorizedAccessException") {
            throw "Access is denied. User does not relevant permissions to create the mini dump of process with ID: {0}" -f $processId
        }
        else {
            throw
        }
    }
}
finally {
    if (Test-Path $fileName) {
        if ((Get-Item $fileName).length -eq 0) {
            # Delete the zero byte file.
            Remove-Item -Path $fileName -Force -ErrorAction Stop
        }
    }
}

}
## [END] New-WACPVProcessDumpDownlevel ##
function Start-WACPVCimProcess {
<#

.SYNOPSIS
Starts new process.

.DESCRIPTION
Starts new process.

.ROLE
Administrators

#>

##SkipCheck=true##

Param(
[string]$CommandLine
)

import-module CimCmdlets

Invoke-CimMethod -Namespace root/Microsoft/Windows/ManagementTools -ClassName MSFT_MTProcess -MethodName CreateProcess -Arguments @{CommandLine=$CommandLine;}

}
## [END] Start-WACPVCimProcess ##
function Start-WACPVProcessDownlevel {
<#

.SYNOPSIS
Start a new process on downlevel computer.

.DESCRIPTION
Start a new process on downlevel computer.

.ROLE
Administrators

#>

param
(
	[Parameter(Mandatory = $true)]
	[string]
	$commandLine
)

Set-StrictMode -Version 5.0

Start-Process $commandLine

}
## [END] Start-WACPVProcessDownlevel ##
function Stop-WACPVCimProcess {
<#

.SYNOPSIS
Stop a process.

.DESCRIPTION
Stop a process.

.ROLE
Administrators

#>

##SkipCheck=true##

Param(
[System.UInt16]$ProcessId
)

import-module CimCmdlets

$instance = New-CimInstance -Namespace root/Microsoft/Windows/ManagementTools -ClassName MSFT_MTProcess -Key @('ProcessId') -Property @{ProcessId=$ProcessId;} -ClientOnly
Remove-CimInstance $instance

}
## [END] Stop-WACPVCimProcess ##
function Stop-WACPVProcesses {
<#

.SYNOPSIS
Stop the process on a computer.

.DESCRIPTION
Stop the process on a computer.

.ROLE
Administrators

#>

param
(
	[Parameter(Mandatory = $true)]
	[int[]]
	$processIds
)

Set-StrictMode -Version 5.0

Stop-Process $processIds -Force

}
## [END] Stop-WACPVProcesses ##
function Get-WACPVCimWin32LogicalDisk {
<#

.SYNOPSIS
Gets Win32_LogicalDisk object.

.DESCRIPTION
Gets Win32_LogicalDisk object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_LogicalDisk

}
## [END] Get-WACPVCimWin32LogicalDisk ##
function Get-WACPVCimWin32NetworkAdapter {
<#

.SYNOPSIS
Gets Win32_NetworkAdapter object.

.DESCRIPTION
Gets Win32_NetworkAdapter object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_NetworkAdapter

}
## [END] Get-WACPVCimWin32NetworkAdapter ##
function Get-WACPVCimWin32PhysicalMemory {
<#

.SYNOPSIS
Gets Win32_PhysicalMemory object.

.DESCRIPTION
Gets Win32_PhysicalMemory object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_PhysicalMemory

}
## [END] Get-WACPVCimWin32PhysicalMemory ##
function Get-WACPVCimWin32Processor {
<#

.SYNOPSIS
Gets Win32_Processor object.

.DESCRIPTION
Gets Win32_Processor object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_Processor

}
## [END] Get-WACPVCimWin32Processor ##
function Get-WACPVClusterInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a cluster.

.DESCRIPTION
Retrieves the inventory data for a cluster.

.ROLE
Readers

#>

Import-Module CimCmdlets -ErrorAction SilentlyContinue

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
Import-Module FailoverClusters -ErrorAction SilentlyContinue

Import-Module Storage -ErrorAction SilentlyContinue
<#

.SYNOPSIS
Get the name of this computer.

.DESCRIPTION
Get the best available name for this computer.  The FQDN is preferred, but when not avaialble
the NetBIOS name will be used instead.

#>

function getComputerName() {
    $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object Name, DNSHostName

    if ($computerSystem) {
        $computerName = $computerSystem.DNSHostName

        if ($null -eq $computerName) {
            $computerName = $computerSystem.Name
        }

        return $computerName
    }

    return $null
}

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell cmdlets installed on this server?

#>

function getIsClusterCmdletAvailable() {
    $cmdlet = Get-Command "Get-Cluster" -ErrorAction SilentlyContinue

    return !!$cmdlet
}

<#

.SYNOPSIS
Get the MSCluster Cluster CIM instance from this server.

.DESCRIPTION
Get the MSCluster Cluster CIM instance from this server.

#>
function getClusterCimInstance() {
    $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue

    if ($namespace) {
        return Get-CimInstance -Namespace root/mscluster MSCluster_Cluster -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object fqdn, S2DEnabled
    }

    return $null
}


<#

.SYNOPSIS
Determines if the current cluster supports Failover Clusters Time Series Database.

.DESCRIPTION
Use the existance of the path value of cmdlet Get-StorageHealthSetting to determine if TSDB
is supported or not.

#>
function getClusterPerformanceHistoryPath() {
    $storageSubsystem = Get-StorageSubSystem clus* -ErrorAction SilentlyContinue
    $storageHealthSettings = Get-StorageHealthSetting -InputObject $storageSubsystem -Name "System.PerformanceHistory.Path" -ErrorAction SilentlyContinue

    return $null -ne $storageHealthSettings
}

<#

.SYNOPSIS
Get some basic information about the cluster from the cluster.

.DESCRIPTION
Get the needed cluster properties from the cluster.

#>
function getClusterInfo() {
    $returnValues = @{}

    $returnValues.Fqdn = $null
    $returnValues.isS2DEnabled = $false
    $returnValues.isTsdbEnabled = $false

    $cluster = getClusterCimInstance
    if ($cluster) {
        $returnValues.Fqdn = $cluster.fqdn
        $isS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -eq 1)
        $returnValues.isS2DEnabled = $isS2dEnabled

        if ($isS2DEnabled) {
            $returnValues.isTsdbEnabled = getClusterPerformanceHistoryPath
        } else {
            $returnValues.isTsdbEnabled = $false
        }
    }

    return $returnValues
}

<#

.SYNOPSIS
Are the cluster PowerShell Health cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell Health cmdlets installed on this server?

s#>
function getisClusterHealthCmdletAvailable() {
    $cmdlet = Get-Command -Name "Get-HealthFault" -ErrorAction SilentlyContinue

    return !!$cmdlet
}
<#

.SYNOPSIS
Are the Britannica (sddc management resources) available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) available on the cluster?

#>
function getIsBritannicaEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual machine available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual machine available on the cluster?

#>
function getIsBritannicaVirtualMachineEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualMachine -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual switch available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual switch available on the cluster?

#>
function getIsBritannicaVirtualSwitchEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualSwitch -ErrorAction SilentlyContinue)
}

###########################################################################
# main()
###########################################################################

$clusterInfo = getClusterInfo

$result = New-Object PSObject

$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $clusterInfo.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2DEnabled' -Value $clusterInfo.isS2DEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsTsdbEnabled' -Value $clusterInfo.isTsdbEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterHealthCmdletAvailable' -Value (getIsClusterHealthCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value (getIsBritannicaEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualMachineEnabled' -Value (getIsBritannicaVirtualMachineEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualSwitchEnabled' -Value (getIsBritannicaVirtualSwitchEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterCmdletAvailable' -Value (getIsClusterCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'CurrentClusterNode' -Value (getComputerName)

$result

}
## [END] Get-WACPVClusterInventory ##
function Get-WACPVClusterNodes {
<#

.SYNOPSIS
Retrieves the inventory data for cluster nodes in a particular cluster.

.DESCRIPTION
Retrieves the inventory data for cluster nodes in a particular cluster.

.ROLE
Readers

#>

import-module CimCmdlets

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
import-module FailoverClusters -ErrorAction SilentlyContinue

###############################################################################
# Constants
###############################################################################

Set-Variable -Name LogName -Option Constant -Value "Microsoft-ServerManagementExperience" -ErrorAction SilentlyContinue
Set-Variable -Name LogSource -Option Constant -Value "SMEScripts" -ErrorAction SilentlyContinue
Set-Variable -Name ScriptName -Option Constant -Value $MyInvocation.ScriptName -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed?

.DESCRIPTION
Use the Get-Command cmdlet to quickly test if the cluster PowerShell cmdlets
are installed on this server.

#>

function getClusterPowerShellSupport() {
    $cmdletInfo = Get-Command 'Get-ClusterNode' -ErrorAction SilentlyContinue

    return $cmdletInfo -and $cmdletInfo.Name -eq "Get-ClusterNode"
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster CIM provider.

.DESCRIPTION
When the cluster PowerShell cmdlets are not available fallback to using
the cluster CIM provider to get the needed information.

#>

function getClusterNodeCimInstances() {
    # Change the WMI property NodeDrainStatus to DrainStatus to match the PS cmdlet output.
    return Get-CimInstance -Namespace root/mscluster MSCluster_Node -ErrorAction SilentlyContinue | `
        Microsoft.PowerShell.Utility\Select-Object @{Name="DrainStatus"; Expression={$_.NodeDrainStatus}}, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster PowerShell cmdlets.

.DESCRIPTION
When the cluster PowerShell cmdlets are available use this preferred function.

#>

function getClusterNodePsInstances() {
    return Get-ClusterNode -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object DrainStatus, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Use DNS services to get the FQDN of the cluster NetBIOS name.

.DESCRIPTION
Use DNS services to get the FQDN of the cluster NetBIOS name.

.Notes
It is encouraged that the caller add their approprate -ErrorAction when
calling this function.

#>

function getClusterNodeFqdn([string]$clusterNodeName) {
    return ([System.Net.Dns]::GetHostEntry($clusterNodeName)).HostName
}

<#

.SYNOPSIS
Writes message to event log as warning.

.DESCRIPTION
Writes message to event log as warning.

#>

function writeToEventLog([string]$message) {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Warning `
        -Message $message  -ErrorAction SilentlyContinue
}

<#

.SYNOPSIS
Get the cluster nodes.

.DESCRIPTION
When the cluster PowerShell cmdlets are available get the information about the cluster nodes
using PowerShell.  When the cmdlets are not available use the Cluster CIM provider.

#>

function getClusterNodes() {
    $isClusterCmdletAvailable = getClusterPowerShellSupport

    if ($isClusterCmdletAvailable) {
        $clusterNodes = getClusterNodePsInstances
    } else {
        $clusterNodes = getClusterNodeCimInstances
    }

    $clusterNodeMap = @{}

    foreach ($clusterNode in $clusterNodes) {
        $clusterNodeName = $clusterNode.Name.ToLower()
        try 
        {
            $clusterNodeFqdn = getClusterNodeFqdn $clusterNodeName -ErrorAction SilentlyContinue
        }
        catch 
        {
            $clusterNodeFqdn = $clusterNodeName
            writeToEventLog "[$ScriptName]: The fqdn for node '$clusterNodeName' could not be obtained. Defaulting to machine name '$clusterNodeName'"
        }

        $clusterNodeResult = New-Object PSObject

        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FullyQualifiedDomainName' -Value $clusterNodeFqdn
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'Name' -Value $clusterNodeName
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DynamicWeight' -Value $clusterNode.DynamicWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'NodeWeight' -Value $clusterNode.NodeWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FaultDomain' -Value $clusterNode.FaultDomain
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'State' -Value $clusterNode.State
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DrainStatus' -Value $clusterNode.DrainStatus

        $clusterNodeMap.Add($clusterNodeName, $clusterNodeResult)
    }

    return $clusterNodeMap
}

###########################################################################
# main()
###########################################################################

getClusterNodes

}
## [END] Get-WACPVClusterNodes ##
function Get-WACPVDecryptedDataFromNode {
<#

.SYNOPSIS
Gets data after decrypting it on a node.

.DESCRIPTION
Decrypts data on node using a cached RSAProvider used during encryption within 3 minutes of encryption and returns the decrypted data.
This script should be imported or copied directly to other scripts, do not send the returned data as an argument to other scripts.

.PARAMETER encryptedData
Encrypted data to be decrypted (String).

.ROLE
Readers

#>
param (
  [Parameter(Mandatory = $true)]
  [String]
  $encryptedData
)

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function DecryptDataWithJWKOnNode {
  if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue) {
    $rsaProvider = (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedData), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
    return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
  }
  # If you copy this script directly to another, you can get rid of the throw statement and add custom error handling logic such as "Write-Error"
  throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
}

}
## [END] Get-WACPVDecryptedDataFromNode ##
function Get-WACPVEncryptionJWKOnNode {
<#

.SYNOPSIS
Gets encrytion JSON web key from node.

.DESCRIPTION
Gets encrytion JSON web key from node.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function Get-RSAProvider
{
    if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue)
    {
        return (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    }

    $Global:RSA = New-Object System.Security.Cryptography.RSACryptoServiceProvider -ArgumentList 4096
    return $RSA
}

function Get-JsonWebKey
{
    $rsaProvider = Get-RSAProvider
    $parameters = $rsaProvider.ExportParameters($false)
    return [PSCustomObject]@{
        kty = 'RSA'
        alg = 'RSA-OAEP'
        e = [Convert]::ToBase64String($parameters.Exponent)
        n = [Convert]::ToBase64String($parameters.Modulus).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }
}

$jwk = Get-JsonWebKey
ConvertTo-Json $jwk -Compress

}
## [END] Get-WACPVEncryptionJWKOnNode ##
function Get-WACPVServerInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a server.

.DESCRIPTION
Retrieves the inventory data for a server.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

Import-Module CimCmdlets

Import-Module Storage -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Converts an arbitrary version string into just 'Major.Minor'

.DESCRIPTION
To make OS version comparisons we only want to compare the major and
minor version.  Build number and/os CSD are not interesting.

#>

function convertOsVersion([string]$osVersion) {
  [Ref]$parsedVersion = $null
  if (![Version]::TryParse($osVersion, $parsedVersion)) {
    return $null
  }

  $version = [Version]$parsedVersion.Value
  return New-Object Version -ArgumentList $version.Major, $version.Minor
}

<#

.SYNOPSIS
Determines if CredSSP is enabled for the current server or client.

.DESCRIPTION
Check the registry value for the CredSSP enabled state.

#>

function isCredSSPEnabled() {
  Set-Variable credSSPServicePath -Option Constant -Value "WSMan:\localhost\Service\Auth\CredSSP"
  Set-Variable credSSPClientPath -Option Constant -Value "WSMan:\localhost\Client\Auth\CredSSP"

  $credSSPServerEnabled = $false;
  $credSSPClientEnabled = $false;

  $credSSPServerService = Get-Item $credSSPServicePath -ErrorAction SilentlyContinue
  if ($credSSPServerService) {
    $credSSPServerEnabled = [System.Convert]::ToBoolean($credSSPServerService.Value)
  }

  $credSSPClientService = Get-Item $credSSPClientPath -ErrorAction SilentlyContinue
  if ($credSSPClientService) {
    $credSSPClientEnabled = [System.Convert]::ToBoolean($credSSPClientService.Value)
  }

  return ($credSSPServerEnabled -or $credSSPClientEnabled)
}

<#

.SYNOPSIS
Determines if the Hyper-V role is installed for the current server or client.

.DESCRIPTION
The Hyper-V role is installed when the VMMS service is available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>

function isHyperVRoleInstalled() {
  $vmmsService = Get-Service -Name "VMMS" -ErrorAction SilentlyContinue

  return $vmmsService -and $vmmsService.Name -eq "VMMS"
}

<#

.SYNOPSIS
Determines if the Hyper-V PowerShell support module is installed for the current server or client.

.DESCRIPTION
The Hyper-V PowerShell support module is installed when the modules cmdlets are available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>
function isHyperVPowerShellSupportInstalled() {
  # quicker way to find the module existence. it doesn't load the module.
  return !!(Get-Module -ListAvailable Hyper-V -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Determines if Windows Management Framework (WMF) 5.0, or higher, is installed for the current server or client.

.DESCRIPTION
Windows Admin Center requires WMF 5 so check the registey for WMF version on Windows versions that are less than
Windows Server 2016.

#>
function isWMF5Installed([string] $operatingSystemVersion) {
  Set-Variable Server2016 -Option Constant -Value (New-Object Version '10.0')   # And Windows 10 client SKUs
  Set-Variable Server2012 -Option Constant -Value (New-Object Version '6.2')

  $version = convertOsVersion $operatingSystemVersion
  if (-not $version) {
    # Since the OS version string is not properly formatted we cannot know the true installed state.
    return $false
  }

  if ($version -ge $Server2016) {
    # It's okay to assume that 2016 and up comes with WMF 5 or higher installed
    return $true
  }
  else {
    if ($version -ge $Server2012) {
      # Windows 2012/2012R2 are supported as long as WMF 5 or higher is installed
      $registryKey = 'HKLM:\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine'
      $registryKeyValue = Get-ItemProperty -Path $registryKey -Name PowerShellVersion -ErrorAction SilentlyContinue

      if ($registryKeyValue -and ($registryKeyValue.PowerShellVersion.Length -ne 0)) {
        $installedWmfVersion = [Version]$registryKeyValue.PowerShellVersion

        if ($installedWmfVersion -ge [Version]'5.0') {
          return $true
        }
      }
    }
  }

  return $false
}

<#

.SYNOPSIS
Determines if the current usser is a system administrator of the current server or client.

.DESCRIPTION
Determines if the current usser is a system administrator of the current server or client.

#>
function isUserAnAdministrator() {
  return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

<#

.SYNOPSIS
Get some basic information about the Failover Cluster that is running on this server.

.DESCRIPTION
Create a basic inventory of the Failover Cluster that may be running in this server.

#>
function getClusterInformation() {
  $returnValues = @{ }

  $returnValues.IsS2dEnabled = $false
  $returnValues.IsCluster = $false
  $returnValues.ClusterFqdn = $null
  $returnValues.IsBritannicaEnabled = $false

  $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue
  if ($namespace) {
    $cluster = Get-CimInstance -Namespace root/MSCluster -ClassName MSCluster_Cluster -ErrorAction SilentlyContinue
    if ($cluster) {
      $returnValues.IsCluster = $true
      $returnValues.ClusterFqdn = $cluster.Fqdn
      $returnValues.IsS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -gt 0)
      $returnValues.IsBritannicaEnabled = $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
    }
  }

  return $returnValues
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

#>
function getComputerFqdnAndAddress($computerName) {
  $hostEntry = [System.Net.Dns]::GetHostEntry($computerName)
  $addressList = @()
  foreach ($item in $hostEntry.AddressList) {
    $address = New-Object PSObject
    $address | Add-Member -MemberType NoteProperty -Name 'IpAddress' -Value $item.ToString()
    $address | Add-Member -MemberType NoteProperty -Name 'AddressFamily' -Value $item.AddressFamily.ToString()
    $addressList += $address
  }

  $result = New-Object PSObject
  $result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $hostEntry.HostName
  $result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $addressList
  return $result
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

#>
function getHostFqdnAndAddress($computerSystem) {
  $computerName = $computerSystem.DNSHostName
  if (!$computerName) {
    $computerName = $computerSystem.Name
  }

  return getComputerFqdnAndAddress $computerName
}

<#

.SYNOPSIS
Are the needed management CIM interfaces available on the current server or client.

.DESCRIPTION
Check for the presence of the required server management CIM interfaces.

#>
function getManagementToolsSupportInformation() {
  $returnValues = @{ }

  $returnValues.ManagementToolsAvailable = $false
  $returnValues.ServerManagerAvailable = $false

  $namespaces = Get-CimInstance -Namespace root/microsoft/windows -ClassName __NAMESPACE -ErrorAction SilentlyContinue

  if ($namespaces) {
    $returnValues.ManagementToolsAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ManagementTools" })
    $returnValues.ServerManagerAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ServerManager" })
  }

  return $returnValues
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>
function isRemoteAppEnabled() {
  Set-Variable key -Option Constant -Value "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Terminal Server\\TSAppAllowList"

  $registryKeyValue = Get-ItemProperty -Path $key -Name fDisabledAllowList -ErrorAction SilentlyContinue

  if (-not $registryKeyValue) {
    return $false
  }
  return $registryKeyValue.fDisabledAllowList -eq 1
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>

<#
c
.SYNOPSIS
Get the Win32_OperatingSystem information as well as current version information from the registry

.DESCRIPTION
Get the Win32_OperatingSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller. Included in the results are current version
information from the registry

#>
function getOperatingSystemInfo() {
  $operatingSystemInfo = Get-CimInstance Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object csName, Caption, OperatingSystemSKU, Version, ProductType, OSType, LastBootUpTime, SerialNumber
  $currentVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" | Microsoft.PowerShell.Utility\Select-Object CurrentBuild, UBR, DisplayVersion

  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name CurrentBuild -Value $currentVersion.CurrentBuild
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name UpdateBuildRevision -Value $currentVersion.UBR
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name DisplayVersion -Value $currentVersion.DisplayVersion

  return $operatingSystemInfo
}

<#

.SYNOPSIS
Get the Win32_ComputerSystem information

.DESCRIPTION
Get the Win32_ComputerSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller.

#>
function getComputerSystemInfo() {
  return Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | `
    Microsoft.PowerShell.Utility\Select-Object TotalPhysicalMemory, DomainRole, Manufacturer, Model, NumberOfLogicalProcessors, Domain, Workgroup, DNSHostName, Name, PartOfDomain, SystemFamily, SystemSKUNumber
}

<#

.SYNOPSIS
Get SMBIOS locally from the passed in machineName


.DESCRIPTION
Get SMBIOS locally from the passed in machine name

#>
function getSmbiosData($computerSystem) {
  <#
    Array of chassis types.
    The following list of ChassisTypes is copied from the latest DMTF SMBIOS specification.
    REF: https://www.dmtf.org/sites/default/files/standards/documents/DSP0134_3.1.1.pdf
  #>
  $ChassisTypes =
  @{
    1  = 'Other'
    2  = 'Unknown'
    3  = 'Desktop'
    4  = 'Low Profile Desktop'
    5  = 'Pizza Box'
    6  = 'Mini Tower'
    7  = 'Tower'
    8  = 'Portable'
    9  = 'Laptop'
    10 = 'Notebook'
    11 = 'Hand Held'
    12 = 'Docking Station'
    13 = 'All in One'
    14 = 'Sub Notebook'
    15 = 'Space-Saving'
    16 = 'Lunch Box'
    17 = 'Main System Chassis'
    18 = 'Expansion Chassis'
    19 = 'SubChassis'
    20 = 'Bus Expansion Chassis'
    21 = 'Peripheral Chassis'
    22 = 'Storage Chassis'
    23 = 'Rack Mount Chassis'
    24 = 'Sealed-Case PC'
    25 = 'Multi-system chassis'
    26 = 'Compact PCI'
    27 = 'Advanced TCA'
    28 = 'Blade'
    29 = 'Blade Enclosure'
    30 = 'Tablet'
    31 = 'Convertible'
    32 = 'Detachable'
    33 = 'IoT Gateway'
    34 = 'Embedded PC'
    35 = 'Mini PC'
    36 = 'Stick PC'
  }

  $list = New-Object System.Collections.ArrayList
  $win32_Bios = Get-CimInstance -class Win32_Bios
  $obj = New-Object -Type PSObject | Microsoft.PowerShell.Utility\Select-Object SerialNumber, Manufacturer, UUID, BaseBoardProduct, ChassisTypes, Chassis, SystemFamily, SystemSKUNumber, SMBIOSAssetTag
  $obj.SerialNumber = $win32_Bios.SerialNumber
  $obj.Manufacturer = $win32_Bios.Manufacturer
  $computerSystemProduct = Get-CimInstance Win32_ComputerSystemProduct
  if ($null -ne $computerSystemProduct) {
    $obj.UUID = $computerSystemProduct.UUID
  }
  $baseboard = Get-CimInstance Win32_BaseBoard
  if ($null -ne $baseboard) {
    $obj.BaseBoardProduct = $baseboard.Product
  }
  $systemEnclosure = Get-CimInstance Win32_SystemEnclosure
  if ($null -ne $systemEnclosure) {
    $obj.SMBIOSAssetTag = $systemEnclosure.SMBIOSAssetTag
  }
  $obj.ChassisTypes = Get-CimInstance Win32_SystemEnclosure | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty ChassisTypes
  $obj.Chassis = New-Object -TypeName 'System.Collections.ArrayList'
  $obj.ChassisTypes | ForEach-Object -Process {
    $obj.Chassis.Add($ChassisTypes[[int]$_])
  }
  $obj.SystemFamily = $computerSystem.SystemFamily
  $obj.SystemSKUNumber = $computerSystem.SystemSKUNumber
  $list.Add($obj) | Out-Null

  return $list

}

<#

.SYNOPSIS
Get the azure arc status information

.DESCRIPTION
Get the azure arc status information

#>
function getAzureArcStatus() {

  $LogName = "Microsoft-ServerManagementExperience"
  $LogSource = "SMEScript"
  $ScriptName = "Get-ServerInventory.ps1 - getAzureArcStatus()"

  Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

  Get-Service -Name himds -ErrorVariable Err -ErrorAction SilentlyContinue | Out-Null

  if (!!$Err) {

    $Err = "The Azure arc agent is not installed. Details: $Err"

    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
    -Message "[$ScriptName]: $Err"  -ErrorAction SilentlyContinue

    $status = "NotInstalled"
  }
  else {
    $status = (azcmagent show --json | ConvertFrom-Json -ErrorAction Stop).status
  }

  return $status
}

<#

.SYNOPSIS
Gets an EnforcementMode that describes the system lockdown policy on this computer.

.DESCRIPTION
By checking the system lockdown policy, we can infer if PowerShell is in ConstrainedLanguage mode as a result of an enforced WDAC policy.
Note: $ExecutionContext.SessionState.LanguageMode should not be used within a trusted (by the WDAC policy) script context for this purpose because
the language mode returned would potentially not reflect the system-wide lockdown policy/language mode outside of the execution context.

#>
function getSystemLockdownPolicy() {
  return [System.Management.Automation.Security.SystemPolicy]::GetSystemLockdownPolicy().ToString()
}

<#

.SYNOPSIS
Determines if the operating system is HCI.

.DESCRIPTION
Using the operating system 'Caption' (which corresponds to the 'ProductName' registry key at HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion) to determine if a server OS is HCI.

#>
function isServerOsHCI([string] $operatingSystemCaption) {
  return $operatingSystemCaption -eq "Microsoft Azure Stack HCI"
}

###########################################################################
# main()
###########################################################################

$operatingSystem = getOperatingSystemInfo
$computerSystem = getComputerSystemInfo
$isAdministrator = isUserAnAdministrator
$fqdnAndAddress = getHostFqdnAndAddress $computerSystem
$hostname = [Environment]::MachineName
$netbios = $env:ComputerName
$managementToolsInformation = getManagementToolsSupportInformation
$isWmfInstalled = isWMF5Installed $operatingSystem.Version
$clusterInformation = getClusterInformation -ErrorAction SilentlyContinue
$isHyperVPowershellInstalled = isHyperVPowerShellSupportInstalled
$isHyperVRoleInstalled = isHyperVRoleInstalled
$isCredSSPEnabled = isCredSSPEnabled
$isRemoteAppEnabled = isRemoteAppEnabled
$smbiosData = getSmbiosData $computerSystem
$azureArcStatus = getAzureArcStatus
$systemLockdownPolicy = getSystemLockdownPolicy
$isHciServer = isServerOsHCI $operatingSystem.Caption

$result = New-Object PSObject
$result | Add-Member -MemberType NoteProperty -Name 'IsAdministrator' -Value $isAdministrator
$result | Add-Member -MemberType NoteProperty -Name 'OperatingSystem' -Value $operatingSystem
$result | Add-Member -MemberType NoteProperty -Name 'ComputerSystem' -Value $computerSystem
$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $fqdnAndAddress.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $fqdnAndAddress.AddressList
$result | Add-Member -MemberType NoteProperty -Name 'Hostname' -Value $hostname
$result | Add-Member -MemberType NoteProperty -Name 'NetBios' -Value $netbios
$result | Add-Member -MemberType NoteProperty -Name 'IsManagementToolsAvailable' -Value $managementToolsInformation.ManagementToolsAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsServerManagerAvailable' -Value $managementToolsInformation.ServerManagerAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsWmfInstalled' -Value $isWmfInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCluster' -Value $clusterInformation.IsCluster
$result | Add-Member -MemberType NoteProperty -Name 'ClusterFqdn' -Value $clusterInformation.ClusterFqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2dEnabled' -Value $clusterInformation.IsS2dEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value $clusterInformation.IsBritannicaEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVRoleInstalled' -Value $isHyperVRoleInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVPowershellInstalled' -Value $isHyperVPowershellInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCredSSPEnabled' -Value $isCredSSPEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsRemoteAppEnabled' -Value $isRemoteAppEnabled
$result | Add-Member -MemberType NoteProperty -Name 'SmbiosData' -Value $smbiosData
$result | Add-Member -MemberType NoteProperty -Name 'AzureArcStatus' -Value $azureArcStatus
$result | Add-Member -MemberType NoteProperty -Name 'SystemLockdownPolicy' -Value $systemLockdownPolicy
$result | Add-Member -MemberType NoteProperty -Name 'IsHciServer' -Value $isHciServer

$result

}
## [END] Get-WACPVServerInventory ##

# SIG # Begin signature block
# MIInvwYJKoZIhvcNAQcCoIInsDCCJ6wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBPj0mK7vaVfL3r
# FJenWbgHYZ5oj4yZ2hgrim1RohMqO6CCDXYwggX0MIID3KADAgECAhMzAAADTrU8
# esGEb+srAAAAAANOMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI5WhcNMjQwMzE0MTg0MzI5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDdCKiNI6IBFWuvJUmf6WdOJqZmIwYs5G7AJD5UbcL6tsC+EBPDbr36pFGo1bsU
# p53nRyFYnncoMg8FK0d8jLlw0lgexDDr7gicf2zOBFWqfv/nSLwzJFNP5W03DF/1
# 1oZ12rSFqGlm+O46cRjTDFBpMRCZZGddZlRBjivby0eI1VgTD1TvAdfBYQe82fhm
# WQkYR/lWmAK+vW/1+bO7jHaxXTNCxLIBW07F8PBjUcwFxxyfbe2mHB4h1L4U0Ofa
# +HX/aREQ7SqYZz59sXM2ySOfvYyIjnqSO80NGBaz5DvzIG88J0+BNhOu2jl6Dfcq
# jYQs1H/PMSQIK6E7lXDXSpXzAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUnMc7Zn/ukKBsBiWkwdNfsN5pdwAw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMDUxNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAD21v9pHoLdBSNlFAjmk
# mx4XxOZAPsVxxXbDyQv1+kGDe9XpgBnT1lXnx7JDpFMKBwAyIwdInmvhK9pGBa31
# TyeL3p7R2s0L8SABPPRJHAEk4NHpBXxHjm4TKjezAbSqqbgsy10Y7KApy+9UrKa2
# kGmsuASsk95PVm5vem7OmTs42vm0BJUU+JPQLg8Y/sdj3TtSfLYYZAaJwTAIgi7d
# hzn5hatLo7Dhz+4T+MrFd+6LUa2U3zr97QwzDthx+RP9/RZnur4inzSQsG5DCVIM
# pA1l2NWEA3KAca0tI2l6hQNYsaKL1kefdfHCrPxEry8onJjyGGv9YKoLv6AOO7Oh
# JEmbQlz/xksYG2N/JSOJ+QqYpGTEuYFYVWain7He6jgb41JbpOGKDdE/b+V2q/gX
# UgFe2gdwTpCDsvh8SMRoq1/BNXcr7iTAU38Vgr83iVtPYmFhZOVM0ULp/kKTVoir
# IpP2KCxT4OekOctt8grYnhJ16QMjmMv5o53hjNFXOxigkQWYzUO+6w50g0FAeFa8
# 5ugCCB6lXEk21FFB1FdIHpjSQf+LP/W2OV/HfhC3uTPgKbRtXo83TZYEudooyZ/A
# Vu08sibZ3MkGOJORLERNwKm2G7oqdOv4Qj8Z0JrGgMzj46NFKAxkLSpE5oHQYP1H
# tPx1lPfD7iNSbJsP6LiUHXH1MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGZ8wghmbAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAANOtTx6wYRv6ysAAAAAA04wDQYJYIZIAWUDBAIB
# BQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEICTif5OLNtNvAgYBWZp+HqYc
# b8YmsWhDTwDSbxhz6HAFMEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEAdmvpo7TsNTi8IG7uDIri48+kd0avXYKb0vM0ahTqRecK75t8tbY9tcaA
# j5YrN73GOe7UygaT+Yn9C/DXA1v+9bemrkFPg7iSJpTXLF1fTVH2xpAd+8BxrKlx
# QYapal+5GL2eTnC64XktD+FvHLMV8K03nhPXQA+ndi0aAUpJq9ZyBLSy8oonqnUl
# 1I+OlcG6y/NC3dgXiEhhHt7LsxTfXn0OQlfjgDOgYsoRzZZiwxG2XsmalSH0EVZK
# vnFOJKPSBFXLSAXauao7p1rgQ02iJ6sI1MtkbXAjKXHra9Y3ci5ciAXmTd47bh/I
# WeCVdxjSKdXSCurDyGC+ZOlqTFh+wKGCFykwghclBgorBgEEAYI3AwMBMYIXFTCC
# FxEGCSqGSIb3DQEHAqCCFwIwghb+AgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFZBgsq
# hkiG9w0BCRABBKCCAUgEggFEMIIBQAIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCB7uQRBRLPqdOvsnClXaWaWfknN2iAri7/SmrDx1hIB3gIGZWdIHKcd
# GBMyMDIzMTIwNjIzNDMxNC42OThaMASAAgH0oIHYpIHVMIHSMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJl
# bGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNO
# OkZDNDEtNEJENC1EMjIwMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNloIIReDCCBycwggUPoAMCAQICEzMAAAHimZmV8dzjIOsAAQAAAeIwDQYJ
# KoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMjMx
# MDEyMTkwNzI1WhcNMjUwMTEwMTkwNzI1WjCB0jELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxhbmQgT3Bl
# cmF0aW9ucyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpGQzQxLTRC
# RDQtRDIyMDElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBALVjtZhV+kFmb8cKQpg2mzis
# DlRI978Gb2amGvbAmCd04JVGeTe/QGzM8KbQrMDol7DC7jS03JkcrPsWi9WpVwsI
# ckRQ8AkX1idBG9HhyCspAavfuvz55khl7brPQx7H99UJbsE3wMmpmJasPWpgF05z
# ZlvpWQDULDcIYyl5lXI4HVZ5N6MSxWO8zwWr4r9xkMmUXs7ICxDJr5a39SSePAJR
# IyznaIc0WzZ6MFcTRzLLNyPBE4KrVv1LFd96FNxAzwnetSePg88EmRezr2T3HTFE
# lneJXyQYd6YQ7eCIc7yllWoY03CEg9ghorp9qUKcBUfFcS4XElf3GSERnlzJsK7s
# /ZGPU4daHT2jWGoYha2QCOmkgjOmBFCqQFFwFmsPrZj4eQszYxq4c4HqPnUu4hT4
# aqpvUZ3qIOXbdyU42pNL93cn0rPTTleOUsOQbgvlRdthFCBepxfb6nbsp3fcZaPB
# fTbtXVa8nLQuMCBqyfsebuqnbwj+lHQfqKpivpyd7KCWACoj78XUwYqy1HyYnStT
# me4T9vK6u2O/KThfROeJHiSg44ymFj+34IcFEhPogaKvNNsTVm4QbqphCyknrwBy
# qorBCLH6bllRtJMJwmu7GRdTQsIx2HMKqphEtpSm1z3ufASdPrgPhsQIRFkHZGui
# hL1Jjj4Lu3CbAmha0lOrAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQURIQOEdq+7Qds
# lptJiCRNpXgJ2gUwHwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYD
# VR0fBFgwVjBUoFKgUIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# cmwvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwG
# CCsGAQUFBwEBBGAwXjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIw
# MjAxMCgxKS5jcnQwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcD
# CDAOBgNVHQ8BAf8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAORURDGrVRTbnulf
# sg2cTsyyh7YXvhVU7NZMkITAQYsFEPVgvSviCylr5ap3ka76Yz0t/6lxuczI6w7t
# Xq8n4WxUUgcj5wAhnNorhnD8ljYqbck37fggYK3+wEwLhP1PGC5tvXK0xYomU1nU
# +lXOy9ZRnShI/HZdFrw2srgtsbWow9OMuADS5lg7okrXa2daCOGnxuaD1IO+65E7
# qv2O0W0sGj7AWdOjNdpexPrspL2KEcOMeJVmkk/O0ganhFzzHAnWjtNWneU11WQ6
# Bxv8OpN1fY9wzQoiycgvOOJM93od55EGeXxfF8bofLVlUE3zIikoSed+8s61NDP+
# x9RMya2mwK/Ys1xdvDlZTHndIKssfmu3vu/a+BFf2uIoycVTvBQpv/drRJD68eo4
# 01mkCRFkmy/+BmQlRrx2rapqAu5k0Nev+iUdBUKmX/iOaKZ75vuQg7hCiBA5xIm5
# ZIXDSlX47wwFar3/BgTwntMq9ra6QRAeS/o/uYWkmvqvE8Aq38QmKgTiBnWSS/uV
# PcaHEyArnyFh5G+qeCGmL44MfEnFEhxc3saPmXhe6MhSgCIGJUZDA7336nQD8fn4
# y6534Lel+LuT5F5bFt0mLwd+H5GxGzObZmm/c3pEWtHv1ug7dS/Dfrcd1sn2E4gk
# 4W1L1jdRBbK9xwkMmwY+CHZeMSvBMIIHcTCCBVmgAwIBAgITMwAAABXF52ueAptJ
# mQAAAAAAFTANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNh
# dGUgQXV0aG9yaXR5IDIwMTAwHhcNMjEwOTMwMTgyMjI1WhcNMzAwOTMwMTgzMjI1
# WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQD
# Ex1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDCCAiIwDQYJKoZIhvcNAQEB
# BQADggIPADCCAgoCggIBAOThpkzntHIhC3miy9ckeb0O1YLT/e6cBwfSqWxOdcjK
# NVf2AX9sSuDivbk+F2Az/1xPx2b3lVNxWuJ+Slr+uDZnhUYjDLWNE893MsAQGOhg
# fWpSg0S3po5GawcU88V29YZQ3MFEyHFcUTE3oAo4bo3t1w/YJlN8OWECesSq/XJp
# rx2rrPY2vjUmZNqYO7oaezOtgFt+jBAcnVL+tuhiJdxqD89d9P6OU8/W7IVWTe/d
# vI2k45GPsjksUZzpcGkNyjYtcI4xyDUoveO0hyTD4MmPfrVUj9z6BVWYbWg7mka9
# 7aSueik3rMvrg0XnRm7KMtXAhjBcTyziYrLNueKNiOSWrAFKu75xqRdbZ2De+JKR
# Hh09/SDPc31BmkZ1zcRfNN0Sidb9pSB9fvzZnkXftnIv231fgLrbqn427DZM9itu
# qBJR6L8FA6PRc6ZNN3SUHDSCD/AQ8rdHGO2n6Jl8P0zbr17C89XYcz1DTsEzOUyO
# ArxCaC4Q6oRRRuLRvWoYWmEBc8pnol7XKHYC4jMYctenIPDC+hIK12NvDMk2ZItb
# oKaDIV1fMHSRlJTYuVD5C4lh8zYGNRiER9vcG9H9stQcxWv2XFJRXRLbJbqvUAV6
# bMURHXLvjflSxIUXk8A8FdsaN8cIFRg/eKtFtvUeh17aj54WcmnGrnu3tz5q4i6t
# AgMBAAGjggHdMIIB2TASBgkrBgEEAYI3FQEEBQIDAQABMCMGCSsGAQQBgjcVAgQW
# BBQqp1L+ZMSavoKRPEY1Kc8Q/y8E7jAdBgNVHQ4EFgQUn6cVXQBeYl2D9OXSZacb
# UzUZ6XIwXAYDVR0gBFUwUzBRBgwrBgEEAYI3TIN9AQEwQTA/BggrBgEFBQcCARYz
# aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9Eb2NzL1JlcG9zaXRvcnku
# aHRtMBMGA1UdJQQMMAoGCCsGAQUFBwMIMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIA
# QwBBMAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFNX2
# VsuP6KJcYmjRPZSQW9fOmhjEMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwu
# bWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8yMDEw
# LTA2LTIzLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYt
# MjMuY3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCdVX38Kq3hLB9nATEkW+Geckv8qW/q
# XBS2Pk5HZHixBpOXPTEztTnXwnE2P9pkbHzQdTltuw8x5MKP+2zRoZQYIu7pZmc6
# U03dmLq2HnjYNi6cqYJWAAOwBb6J6Gngugnue99qb74py27YP0h1AdkY3m2CDPVt
# I1TkeFN1JFe53Z/zjj3G82jfZfakVqr3lbYoVSfQJL1AoL8ZthISEV09J+BAljis
# 9/kpicO8F7BUhUKz/AyeixmJ5/ALaoHCgRlCGVJ1ijbCHcNhcy4sa3tuPywJeBTp
# kbKpW99Jo3QMvOyRgNI95ko+ZjtPu4b6MhrZlvSP9pEB9s7GdP32THJvEKt1MMU0
# sHrYUP4KWN1APMdUbZ1jdEgssU5HLcEUBHG/ZPkkvnNtyo4JvbMBV0lUZNlz138e
# W0QBjloZkWsNn6Qo3GcZKCS6OEuabvshVGtqRRFHqfG3rsjoiV5PndLQTHa1V1QJ
# sWkBRH58oWFsc/4Ku+xBZj1p/cvBQUl+fpO+y/g75LcVv7TOPqUxUYS8vwLBgqJ7
# Fx0ViY1w/ue10CgaiQuPNtq6TPmb/wrpNPgkNWcr4A245oyZ1uEi6vAnQj0llOZ0
# dFtq0Z4+7X6gMTN9vMvpe784cETRkPHIqzqKOghif9lwY1NNje6CbaUFEMFxBmoQ
# tB1VM1izoXBm8qGCAtQwggI9AgEBMIIBAKGB2KSB1TCB0jELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxh
# bmQgT3BlcmF0aW9ucyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpG
# QzQxLTRCRDQtRDIyMDElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vy
# dmljZaIjCgEBMAcGBSsOAwIaAxUAFpuZafp0bnpJdIhfiB1d8pTohm+ggYMwgYCk
# fjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQD
# Ex1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQUFAAIF
# AOkbAIkwIhgPMjAyMzEyMDYyMjE1MzdaGA8yMDIzMTIwNzIyMTUzN1owdDA6Bgor
# BgEEAYRZCgQBMSwwKjAKAgUA6RsAiQIBADAHAgEAAgIVFDAHAgEAAgISrTAKAgUA
# 6RxSCQIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAID
# B6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBBQUAA4GBANKE5Ug76NoOf7nf0qwd
# H+jrJwpSCfFH9/UAZ7WHRuoJKhjvrVMY4K1oRLUraAhHTMFz1JrX4fNvcIPylM72
# Fd8CySEZbu0Ep8X8z1XVRQykB2NqqNhnz/lVfQtv45aOTu1VficriP0exByaGPRf
# QBxcHXrKH7RLJ42DUa/WcChwMYIEDTCCBAkCAQEwgZMwfDELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBIDIwMTACEzMAAAHimZmV8dzjIOsAAQAAAeIwDQYJYIZIAWUDBAIB
# BQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQx
# IgQgTxF+320j8JsT/0KZe6IMTG5BmFLYxziQLLmHh1KTeNkwgfoGCyqGSIb3DQEJ
# EAIvMYHqMIHnMIHkMIG9BCAriSpKEP0muMbBUETODoL4d5LU6I/bjucIZkOJCI9/
# /zCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAB4pmZ
# lfHc4yDrAAEAAAHiMCIEIOGwHA26zH++d19mY55O+ru2/tEv0nh4IzMh1hMm493S
# MA0GCSqGSIb3DQEBCwUABIICAFYwNV8plOn2HLiWp7Os9FL24OcaN5Z8/p5gqSh8
# ozDnikTPAHwRHPep0yRcQM2urIltq9buH5pyIvNROSifZdZI3If5bqIQakRFHelA
# MqED2xY8qJH+ExKGmDee7xd78GJ2WkpYPEvU0gVw6SnnNKLfpxZI+FtQT6m0UX0V
# KlsJ0/P+jRc8M577GAHZ4Ns4A6rTMW8qu8NAwk5GwlUxTvcbyyazT3WcX6UGwYir
# iZSJg1KkWjW/VHOXBRmT7AdlI4tZQQ7n7j5prWSAw3Tzcl30S6n6R5VHIMBJtgcj
# kZoPqyH+p0blDTUzMUwK8/7us9StgBIqqGql22x48vtArOhrCHw19IKIazODamI6
# TRq4573nw+avL0bmV7a2IXkuuTL8v3Fr2NzVWWFpc3ssXie7Uc7oN7wZXSvUoDCy
# BR54RmV5fKAc+lQV4TE+FxNUlWjsVNmY0DE9Kgv1k0FkGNt29Zf2Jd7YoBxaZtMa
# 4OpNJh7f31Ia8C9s4r2ffC9vj6eaCsLSJSM2/lLRhbI4cRjzsShE9y+Qk6sT+cIx
# PrcIpSOQdYXM7yQr4iBDHQlGqP2+rxTN0K+btPBBuBh+fssRftXetnVanfNHMH4h
# ubiorHW4F+ah3mFpNAqyKCa+LKN/E/cMjS3ufLLw8pCYZ6xZNMr/ENaQ8rTSAPb/
# QHly
# SIG # End signature block
