function Disable-ProtectedEventLogging 
{  
    Remove-Item HKLM:\Software\Policies\Microsoft\Windows\EventLog\ProtectedEventLogging -Force –Recurse 
}

function Disable-PSScriptBlockLogging 
{  
    Remove-Item HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging -Force –Recurse 
}

function Disable-PSTranscription 
{  
    Remove-Item HKLM:\Software\Policies\Microsoft\Windows\PowerShell\Transcription -Force –Recurse 
}



Disable-ProtectedEventLogging
Disable-PSScriptBlockLogging 
Disable-PSTranscription