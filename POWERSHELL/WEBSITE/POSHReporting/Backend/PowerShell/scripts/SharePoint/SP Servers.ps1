Get-SPServer | Select @{N="Server Name";E={$_.Address}}, 
                      Role, 
                      @{N="State";E={$_.Status}},
                      @{N="Can Upgrade";E={$_.CanUpgrade}},@{N="Needs Upgrade";E={$_.NeedsUpgrade}} 