$CtrxSvrs = (Find-ADComputer misctrx).name
$ErrorAction = "SilentlyContinue"

foreach ( $Server in $CtrxSvrs )
    {
    Write-Host $Server
    $Users = Get-ChildItem "//$Server/C$/Users/"
    foreach ( $User in $Users )
        {
        $UserName = $User.Name
        If ( Get-ADUser -Filter { SAMAccountName -eq $UserName } ) 
            {
            Write-Host "$($UserName) Exists" -ForegroundColor White
            }
        Else
            {
            Write-Host "$($UserName) Does not exist. Delete Profile on $($Server)? [y,N]" -ForegroundColor Yellow
            $Answer = Read-Host
            If ( $Answer -eq "y" )
                {
                Foreach ( $SubSrv in $CtrxSvrs )
                    {
                    $Profile = Get-WMIObject Win32_UserProfile -ComputerName $SubSrv | ? LocalPath -Match $UserName
                    If ( $Profile )
                        {
                        Write-Host "Deleting $($UserName) on $SubSrv"
                        $Profile.Delete()
                        }
                    }
                }
            }
        }
    }
