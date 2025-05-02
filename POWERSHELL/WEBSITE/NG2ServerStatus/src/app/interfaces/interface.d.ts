export interface IServer {
    SysName: string;
    Disk: DiskInfo[];
    CPU: CpuInfo[];
    Computer: ComputerInfo;
    OS: OSInfo;
    Network: NetworkInfo[];
    Memory: MemoryInfo[];
    Reboots: RebootsInfo[];
    Errors: ErrorsInfo[];
    Warnings: WarningsInfo[];
    Warnings24h: number;
    Errors24h: number;
    Reboots24h: number;
    Warnings7d: number;
    Errors7d: number;
    Reboots7d: number;
    Generated: string;
}

export interface DiskInfo {
    DeviceID: string;
    FileSystem: string;
    Size: number;
    FreeSpace: number;
    VolumeName: string;
    VolumeDirty: boolean;
}

export interface CpuInfo {
    Name: string;
    Manufacturer: string;
    MaxClockSpeed: string;
    LoadPercentage: number;
    NumberOfCores: number;
    NumberOfLogicalProcessors: number;
}

export interface ComputerInfo {
    Caption: string;
    Domain: string;
    Model: string;
    Name: string;
    NumberOfProcessors: number;
}

export interface OSInfo {
    BuildNumber: string;
    Caption: string;
    CSDVersion: string;
    Version: string;
    OSArchitecture: string;
    TotalVisibleMemorySize: number;
    FreePhysicalMemory: number;
    TotalVirtualMemorySize: number;
    FreeVirtualMemory: number;
    OSType: number;
    LastBootUpTime: string;
    Uptime: string;
}

export interface IPAddress {
    IP: string;
}

export interface NetworkInfo {
    Description: string;
    DNSDomain: string;
    DHCPServer: string;
    IPAddress: IPAddress[];
    MACAddress: string;
}

export interface MemoryInfo {
    Capacity: number;
    BankLabel: string;
    DeviceLocator: string;
    PartNumber: string;
    SerialNumber: string;
}

export interface EntryType {
    value: number;
    Value: string;
}

export interface RebootsInfo {
    EventID: number;
    Message: string;
    Source: string;
    TimeGenerated: string;
    EntryType: EntryType;
    TimeWritten: string;
    Log: string;
}

export interface WarningsInfo {
    EventID: number;
    Message: string;
    Source: string;
    TimeGenerated: string;
    EntryType: EntryType;
    TimeWritten: string;
    Log: string;
}

export interface ErrorsInfo {
    EventID: number;
    Message: string;
    Source: string;
    TimeGenerated: string;
    EntryType: EntryType;
    TimeWritten: string;
    Log: string;
}

// Ping result

export interface PingResult {
    Domain: string;
    Alive: string[];
    Dead: string[];
}
