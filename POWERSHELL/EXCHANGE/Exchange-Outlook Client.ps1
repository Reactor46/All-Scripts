    Get-ChildItem -Path '\\lasexdb02\D$\Program Files\Microsoft\Exchange Server\V14\Logging\RPC Client Access\*.log' |
    Get-MrRCAProtocolLog | Out-GridView -Title 'Outlook Client Versions' #|Select-Object -Property Name -Unique | Sort-Object -Property Name -Descending |
    


# Where-Object LastWriteTime -gt (Get-Date).AddDays(-360) 
#Select-Object -Property Version -Unique |
#Sort-Object -Property Version -Descending
    
    
    




