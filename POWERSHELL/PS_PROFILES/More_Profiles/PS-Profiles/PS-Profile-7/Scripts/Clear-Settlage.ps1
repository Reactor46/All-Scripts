$CtrxSvrs = (Find-ADComputer misctrx).name
$UserName = "lsettlage"
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
