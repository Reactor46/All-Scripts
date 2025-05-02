@{
    'Rules' = @{
        'PSAvoidUsingCmdletAliases' = @{
            'Whitelist' = @(
                "ls", "cd", "gsp", "where", "%", "?",
                "foreach", "select", "copy")
        }
    }
}
