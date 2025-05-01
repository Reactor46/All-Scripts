Register-ArgumentCompleter -CommandName powershell -ScriptBlock {
	param ([string]$wordToComplete, $commandAst, $cursorPosition)
	function COMPGEN($txt, $tip) { New-Object System.Management.Automation.CompletionResult -ArgumentList $txt, $txt, 'ParameterName', $tip }
	@(
		COMPGEN '-Command' 'Executes the specified commands (and any parameters) as though they were typed at the Windows PowerShell command prompt, and then exits, unless NoExit is specified. The value of Command can be "-", a string. or a script block.'
		COMPGEN '-ConfigurationName' 'Specifies a configuration endpoint in which Windows PowerShell is run. This can be any endpoint registered on the local machine including the default Windows PowerShell remoting endpoints or a custom endpoint having specific user role capabilities.'
		COMPGEN '-EncodedCommand' 'Accepts a base-64-encoded string version of a command. Use this parameter to submit commands to Windows PowerShell that require complex quotation marks or curly braces.'
		COMPGEN '-ExecutionPolicy' 'Sets the default execution policy for the current session and saves it in the $env:PSExecutionPolicyPreference environment variable. This parameter does not change the Windows PowerShell execution policy that is set in the registry.'
		COMPGEN '-File' 'Runs the specified script in the local scope ("dot-sourced"), so that the functions and variables that the script creates are available in the current session. Enter the script file path and any parameters. File must be the last parameter in the command, because all characters typed after the File parameter name are interpreted as the script file path followed by the script parameters.'
		COMPGEN '-InputFormat' 'Describes the format of data sent to Windows PowerShell. Valid values are "Text" (text strings) or "XML" (serialized CLIXML format).'
		COMPGEN '-Mta' 'Start the shell using a multithreaded apartment.'
		COMPGEN '-NoExit' 'Does not exit after running startup commands.'
		COMPGEN '-NoLogo' 'Hides the copyright banner at startup.'
		COMPGEN '-NonInteractive' 'Does not present an interactive prompt to the user.'
		COMPGEN '-NoProfile' 'Does not load the Windows PowerShell profile.'
		COMPGEN '-OutputFormat' 'Determines how output from Windows PowerShell is formatted. Valid values are "Text" (text strings) or "XML" (serialized CLIXML format).'
		COMPGEN '-PSConsoleFile' 'Loads the specified Windows PowerShell console file. To create a console file, use Export-Console in Windows PowerShell.'
		COMPGEN '-Sta' 'Starts the shell using a single-threaded apartment. Single-threaded apartment (STA) is the default.'
		COMPGEN '-Version' 'Starts the specified version of Windows PowerShell. Enter a version number with the parameter, such as "-version 2.0".'
		COMPGEN '-WindowStyle' 'Sets the window style to Normal, Minimized, Maximized or Hidden.'
	) | Where-Object CompletionText -Like "$wordToComplete*"
}

Register-ArgumentCompleter -CommandName code -Native -ScriptBlock {
	param ([string]$wordToComplete, $commandAst, $cursorPosition)

	(Get-ChildItem $HOME\github.com) | ? Name -Like "$wordToComplete*" | % FullName
	(Get-ChildItem -Name) -like "$wordToComplete*" | Resolve-Path -Relative
}

Register-ArgumentCompleter -Native -CommandName dotnet -ScriptBlock {
	param ($commandName, $wordToComplete, $cursorPosition)
	dotnet complete --position $cursorPosition "$wordToComplete" | ForEach-Object {
		[System.Management.Automation.CompletionResult]::new($_, $_, 'ParameterValue', $_)
	}
}
