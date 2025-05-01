Set-StrictMode -Version Latest
$GLOBAL:ErrorActionPreference = "Inquire"
$GLOBAL:DebugPreference = "Inquire"

$h1 = $ENV:USERPROFILE
$m1 = "$h1\powershell\pShellMain.psm1"

$GLOBAL:ACTIVATE_PROMPT = $null
if ( test-path $m1 -PathType leaf ) {
#   $GLOBAL:profilem = $m1
#   import-Module $m1
}
#$GLOBAL:ACTIVATE_PROMPT = $true
Set-StrictMode -Version Latest
$GLOBAL:ErrorActionPreference = "Inquire"   # Other setting == "SilentlyContinue"
$GLOBAL:DebugPreference = "Inquire"         # Other setting == "SilentlyContinue"
#------------------------------------------------------------------------------
# MAIN powershell profile module
#
# FOR nomad/triad
#
# Run...
#
# Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Remotesigned
#
# ...if powershell startup errors out of this script on a
# permission error. You do not need administrator priv to run...
#
#           set-executionpolicy -scope currentuser
#
# ...will solve the problem
#
#-----------------------------------------------------------------------------
#
# This module will take the place of $profile if
# you place the following lines in $profile ...
#
#           Set-StrictMode -Version 2
#
#           $d  = "\\$domainUsrDir"
#           $n  = $env:username
#           $m  = "$d\$n\powershell\pShellMain.psml"
#
#           if ( test-path $m -PathType leaf ) {
#               $GLOBAL:profilem = $m
#               import-Module $m
#           }
#
# ...the 'GLOBAL:' prefixes really are necessary where they are used.
# The advantage of using a module is that you are not at the mercy
# of Microsoft as to where the module lives. otoh your posh
# startup must be in the directory that $GLOBAL:profile resolves to.
# That directory probably has a space in its name. Enough said.
#
#-----------------------------------------------------------------------------
#
# =========================================================================
# \\$domainUsrDir\$uid\powershell\glblvbl.psml {{{
#
$GLOBAL:maximumHistorycount             =  32767
$GLOBAL:MaximumVariableCount            =  16383
$GLOBAL:FormatEnumerationLimit          =  64
#
# It is best never to overlook the importance of ASCII Codes
#
$GLOBAL:nullstr                         =  ""
$GLOBAL:dash                            =  [string][char][int]0x2d
$GLOBAL:slash                           =  [string][char][int]0x2f
$GLOBAL:backslash                       =  [string][char][int]0x5c
$GLOBAL:singlequote                     =  [string][char][int]0x27
$GLOBAL:doublequote                     =  [string][char][int]0x22
$GLOBAL:newline                         =  [string][char][int]0x0a
$GLOBAL:cReturn                         =  [string][char][int]0x0d
$GLOBAL:hTab                            =  [string][char][int]0x09
$GLOBAL:space                           =  [string][char][int]0x20
$GLOBAL:dollar                          =  [string][char][int]0x24
$GLOBAL:NAK                             =  [string][char][int]0x15
$GLOBAL:EOT                             =  [string][char][int]0x04
#
# there still remain some operations that are easier in COM than anywhere else
#
$GLOBAL:fso                          =   New-Object -com scripting.filesystemobject
$GLOBAL:wshell                       =   New-Object -com WScript.shell
$GLOBAL:pVerifiedS                   =   $null
$GLOBAL:pVerifiedL                   =   $null
#
# More explicit global defs are required in a module than in...
#
# \\$domainUsrDir\$uid\My Documents\windowsPowerShell\Microsoft.PowerShell_profile.ps1
#
# ...I don't like it. But, I think I`m gonna have to live with it.
#
$GLOBAL:GVtsprefix                    = $nullstr
$GLOBAL:uid                           = $env:username
$GLOBAL:folderHash                    = @{}
$GLOBAL:pathCache                     = @{}
$GLOBAL:diary                         = @()
$GLOBAL:abw                           = $null
$GLOBAL:abwL                          = $null
$GLOBAL:bin                           = $null
$GLOBAL:binL                          = $null
$GLOBAL:sqlite                        = $null
$GLOBAL:sqliteL                       = $null
$GLOBAL:bjPath                        = $null
$GLOBAL:bzipexe                       = $null
$GLOBAL:cdJournal                     = $null
$GLOBAL:cdTrail                       = $null
$GLOBAL:psDiaryLclD                   = $null
$GLOBAL:psDiaryTemp                   = $null
$GLOBAL:psDiaryNetD                   = $null
$GLOBAL:psDiaryArchive                = $null
$GLOBAL:psDiaryAnathema               = $null
$GLOBAL:psDiaryTwins                  = $null
$GLOBAL:psDiaryRecent                 = $null
$GLOBAL:psDiaryXmlFile                = $null
$GLOBAL:psDiaryTxtFile                = $null
$GLOBAL:psChronicleFile               = $null
$GLOBAL:ps1Dir                        = $null
$GLOBAL:psm1Dir                       = $null
$GLOBAL:sundryLclD                    = $null
$GLOBAL:sundryNetD                    = $null
$GLOBAL:sundryTmpD                    = $null
$GLOBAL:consoleTitleMode              = $null
$GLOBAL:crunchyvim                    = $null
$GLOBAL:crunchyvimexe                 = $null
$GLOBAL:din                           = $null
$GLOBAL:dinL                          = $null
$GLOBAL:docset                        = $null
$GLOBAL:docsetL                       = $null
$GLOBAL:dyd                           = $null
$GLOBAL:dydL                          = $null
$GLOBAL:eclipseDir                    = $null
$GLOBAL:eclipseimage                  = $null
$GLOBAL:eclipsevm                     = $null
$GLOBAL:eclipsevmexe                  = $null
$GLOBAL:expiredFiles                  = $null
$GLOBAL:ekd                           = $null
$GLOBAL:ekdL                          = $null
$GLOBAL:fileServer                    = $null
$GLOBAL:getLocationRE                 = $null
$GLOBAL:getLongNameRE                 = $null
$GLOBAL:getPSPathRE                   = $null
$GLOBAL:gitRoot                       = $null
$GLOBAL:gitRootL                      = $null
$GLOBAL:git                           = $null
$GLOBAL:gitL                          = $null
$GLOBAL:gitubin                       = $null
$GLOBAL:gitubinL                      = $null
$GLOBAL:gooeyvim                      = $null
$GLOBAL:gooeyvimexe                   = $null
$GLOBAL:hDriveMountPt                 = $null
$GLOBAL:hostfile                      = $null
$GLOBAL:homeLcl                       = $null
$GLOBAL:homeLclDsk                    = $null
$GLOBAL:homeLclRoot                   = $null
$GLOBAL:homeNet                       = $null
$GLOBAL:include                       = $null
$GLOBAL:includeL                      = $null
$GLOBAL:initialDirectory              = $null
$GLOBAL:etcDirectory                  = $null
$GLOBAL:j7d                           = $null
$GLOBAL:j7dL                          = $null
$GLOBAL:lib                           = $null
$GLOBAL:libL                          = $null
$GLOBAL:libpath                       = $null
$GLOBAL:libpathL                      = $null
$GLOBAL:lmportrt0Path                 = $null
$GLOBAL:localhost                     = $null
$GLOBAL:mmm                           = $null
$GLOBAL:mgw                           = $null
$GLOBAL:mgwL                          = $null
$GLOBAL:netLcl                        = $null
$GLOBAL:netLclDsk                     = $null
$GLOBAL:notAfileSystemRE              = $null
$GLOBAL:oslcPath                      = $null
$GLOBAL:path                          = $null
$GLOBAL:pathL                         = $null
$GLOBAL:PERL5LIB                      = $null
$GLOBAL:PERL5LIBL                     = $null
$GLOBAL:PYTHONPATH                    = $null
$GLOBAL:PYTHONPATHL                   = $null
$GLOBAL:pfd                           = $null
$GLOBAL:pfdL                          = $null
$GLOBAL:pfx                           = $null
$GLOBAL:pfxL                          = $null
$GLOBAL:clPath                        = $null
$GLOBAL:pkPath                        = $null
$GLOBAL:psprofilBkLclD                = $null
$GLOBAL:psprofilBkNetD                = $null
$GLOBAL:pyd                           = $null
$GLOBAL:pydL                          = $null
$GLOBAL:rt1DevPath                    = $null
$GLOBAL:rt1ForceActive                = $null
$GLOBAL:rt1Path                       = $null
$GLOBAL:sambahost                     = $null
$GLOBAL:sby                           = $null
$GLOBAL:sbyL                          = $null
$GLOBAL:spc                           = $null
$GLOBAL:spcL                          = $null
$GLOBAL:sqlCmdDir                     = $null
$GLOBAL:sys32                         = $null
$GLOBAL:sys32L                        = $null
$GLOBAL:sys64                         = $null
$GLOBAL:sys64L                        = $null
$GLOBAL:sysInternals                  = $null
$GLOBAL:transcriptdir                 = $null
$GLOBAL:tStampRE                      = $null
$GLOBAL:tStemplate0                   = $null
$GLOBAL:tStemplate1                   = $null
$GLOBAL:tulavnecot                    = $null
$GLOBAL:vim                           = $null
$GLOBAL:vimL                          = $null
$GLOBAL:vimBk                         = $null
$GLOBAL:vimBkL                        = $null
$GLOBAL:vimBackupDir                  = $null
$GLOBAL:vimBackupDirL                 = $null
$GLOBAL:vimLcl                        = $null
$GLOBAL:vimLclL                       = $null
$GLOBAL:vimprofile                    = $null
$GLOBAL:vimprofileL                   = $null
$GLOBAL:vimProfileDir                 = $null
$GLOBAL:vimProfileLclD                = $null
$GLOBAL:vimProfileNetD                = $null
$GLOBAL:vimprofold                    = $null
$GLOBAL:vimrc_base_name               = $null
$GLOBAL:vimrc_base_old                = $null
$GLOBAL:vimruntime                    = $null
$GLOBAL:vimVersion                    = $null
$GLOBAL:vvvIsVerbose                  = $null
$GLOBAL:win                           = $null
$GLOBAL:winL                          = $null
$GLOBAL:dbx                           = $null
$GLOBAL:dbxL                          = $null
$GLOBAL:loh                           = $null
$GLOBAL:lohL                          = $null
$GLOBAL:JAVA_HOME                     = $null
$GLOBAL:JRE_HOME                      = $null
$GLOBAL:CLASSPATH                     = $null
$GLOBAL:JRE8_64                       = $null
$GLOBAL:ECLIPSE_HOME                  = $null
$GLOBAL:SMARTGIT_JAVA_HOME            = $null
$GLOBAL:adb                           = $null
$GLOBAL:xbn                           = $null
$GLOBAL:adhL                          = $null
$GLOBAL:doesNotExist                  = $null
$GLOBAL:doesNotExistL                 = $null
$GLOBAL:oracle                        = $null
$GLOBAL:oracleL                       = $null
$GLOBAL:droid                         = $null
$GLOBAL:droidL                        = $null
#
#
# set up regular expressions and substitutions to manage
# space embeded pathnames from the command-line with a bit
# more convenience
#
$GLOBAL:uComma                          = [string][char][int] 0x002C
$GLOBAL:uComma_re                       = [regex] '(?six: \u002C )'
$GLOBAL:uCommaAKA                       = [string][char][int] 0x25BC
$GLOBAL:uCommaAKA_re                    = [regex] '(?six: \u25BC )'

$GLOBAL:uSQuote                         = [string][char][int] 0x0027
$GLOBAL:uSQuote_re                      = [regex] '(?six: \u0027 )'
$GLOBAL:uSQuoteAKA                      = [string][char][int] 0x25B2
$GLOBAL:uSQuoteAKA_re                   = [regex] '(?six: \u25B2 )'

$GLOBAL:uSpace                          = [string][char][int] 0x0020
$GLOBAL:uSpace_re                       = [regex] '(?six: \u0020 )'
$GLOBAL:uSpaceAKA                       = [string][char][int] 0x2022
$GLOBAL:uSpaceAKA_re                    = [regex] '(?six: \u2022 )'

$GLOBAL:uLparen                         = [string][char][int] 0x0028
$GLOBAL:uLparen_re                      = [regex] '(?six: \u0028 )'
$GLOBAL:uLparenAKA                      = [string][char][int] 0x2039
$GLOBAL:uLparenAKA_re                   = [regex] '(?six: \u2039 )'

$GLOBAL:uRparen                         = [string][char][int] 0x0029
$GLOBAL:uRparen_re                      = [regex] '(?six: \u0029 )'
$GLOBAL:uRparenAKA                      = [string][char][int] 0x203A
$GLOBAL:uRparenAKA_re                   = [regex] '(?six: \u203A )'

$GLOBAL:uLsBrak                         = [string][char][int] 0x005B
$GLOBAL:uLsBrak_re                      = [regex] '(?six: \u005B )'
$GLOBAL:uLsBrakAKA                      = [string][char][int] 0x2264
$GLOBAL:uLsBrakAKA_re                   = [regex] '(?six: \u2264 )'

$GLOBAL:uRsBrak                         = [string][char][int] 0x005D
$GLOBAL:uRsBrak_re                      = [regex] '(?six: \u005D )'
$GLOBAL:uRsBrakAKA                      = [string][char][int] 0x2265
$GLOBAL:uRsBrakAKA_re                   = [regex] '(?six: \u2265 )'

$GLOBAL:uLcBrak                         = [string][char][int] 0x007B
$GLOBAL:uLcBrak_re                      = [regex] '(?six: \u007B )'
$GLOBAL:uLcBrakAKA                      = [string][char][int] 0x25C4
$GLOBAL:uLcBrakAKA_re                   = [regex] '(?six: \u25C4 )'

$GLOBAL:uRcBrak                         = [string][char][int] 0x007D
$GLOBAL:uRcBrak_re                      = [regex] '(?six: \u007D )'
$GLOBAL:uRcBrakAKA                      = [string][char][int] 0x25BA
$GLOBAL:uRcBrakAKA_re                   = [regex] '(?six: \u25BA )'

$GLOBAL:uPound                         = [string][char][int] 0x0023
$GLOBAL:uPound_re                      = [regex] '(?six: \u0023 )'
$GLOBAL:uPoundAKA                      = [string][char][int] 0x220F
$GLOBAL:uPoundAKA_re                   = [regex] '(?six: \u220F )'

#
# These regex transforms could be hard coded ( as in deSpaceEZ ).
# There are only 10. But then I would not have a working example
# of a posh loop through a 2 dimensional array. This, more general,
# solution will make it easy to move other troublesome characters
# out of the way.
#
$GLOBAL:chrAkaXform = (
    ( $uSpace_re,   $uSpaceAKA,   $uSpaceAKA_re,  $uSpace    ),
    ( $uComma_re,   $uCommaAKA,   $uCommaAKA_re,  $uComma    ),
    ( $uSQuote_re,  $uSQuoteAKA,  $uSQuoteAKA_re, $uSQuote   ),
    ( $uLparen_re,  $uLparenAKA,  $uLparenAKA_re, $uLparen   ),
    ( $uRparen_re,  $uRparenAKA,  $uRparenAKA_re, $uRparen   ),
    ( $uLsBrak_re,  $uLsBrakAKA,  $uLsBrakAKA_re, $uLsBrak   ),
    ( $uRsBrak_re,  $uRsBrakAKA,  $uRsBrakAKA_re, $uRsBrak   ),
    ( $uLcBrak_re,  $uLcBrakAKA,  $uLcBrakAKA_re, $uLcBrak   ),
    ( $uRcBrak_re,  $uRcBrakAKA,  $uRcBrakAKA_re, $uRcBrak   ),
    ( $uPound_re,   $uPoundAKA,   $uPoundAKA_re,  $uPound    )
)

$GLOBAL:regx2aka = @"
    (?six:  $uSpace_re      |
            $uComma_re      |
            $uSQuote_re     |
            $uLparen_re     |
            $uRparen_re     |
            $uLsBrak_re     |
            $uRsBrak_re     |
            $uLcBrak_re     |
            $uRcBrak_re     |
            $uPound_re      )

"@
$GLOBAL:regx2aka = [regex]$regx2aka.trim()

$GLOBAL:regx2org = @"
    (?six:  $uSpaceAKA_re   |
            $uCommaAKA_re   |
            $uSQuoteAKA_re  |
            $uLparenAKA_re  |
            $uRparenAKA_re  |
            $uLsBrakAKA_re  |
            $uRsBrakAKA_re  |
            $uLcBrakAKA_re  |
            $uRcBrakAKA_re  |
            $uPoundAKA_re   )
"@
$GLOBAL:regx2org = [regex]$regx2org.trim()

$GLOBAL:sowTypeAllowed = @"
    (?six: \A
    (System[.]Array)                      |
    (System[.]String)                     |
    (System[.]Collections[.]ArrayList) \z )
"@
$GLOBAL:sowTypeAllowed = [regex]$sowTypeAllowed.trim()

$GLOBAL:LeadingWS = [regex] '(?six:  \A  \s+)'
$GLOBAL:TrailingWS = [regex] '(?six: \s+ \z )'
$GLOBAL:hostfile = "C:/Windows/system32/drivers/etc/hosts"

$GLOBAL:CL_TEMPLATE = @"
    <Objs Version="1.1.0.1" xmlns="http://schemas.microsoft.com/powershell/2004/04">
    <Obj RefId="0">
    <TN RefId="0">
    <T>Microsoft.PowerShell.Commands.HistoryInfo</T>
    <T>System.Object</T>
    </TN>
    <ToString>write-host NNNNNNN-NEVER-CONFUSE-MOTION-WITH-ACTION</ToString>
    <Props>
    <I64 N="Id">01</I64>
    <S N="CommandLine">write-host NNNNNNN-NEVER-CONFUSE-MOTION-WITH-ACTION</S>
    <Obj N="ExecutionStatus" RefId="1">
    <TN RefId="1">
    <T>System.Management.Automation.Runspaces.PipelineState</T>
    <T>System.Enum</T>
    <T>System.ValueType</T>
    <T>System.Object</T>
    </TN>
    <ToString>Completed</ToString>
    <I32>4</I32>
    </Obj>
    <DT N="StartExecutionTime">1989-11-09T22:45:00.NNNNNNN-00:00</DT>
    <DT N="EndExecutionTime">1989-11-09T22:45:00.NNNNNNN-00:00</DT>
    </Props>
    </Obj>
    </Objs>
"@

#
# }}}
# =========================================================================
#
# =========================================================================
# \\$domainUsrDir\$uid\powershell\akalist.psml {{{
#
Set-Alias   -scope global -name   eee           -value   "C:/PROGRA~1/SciTE/SciTE.exe"
Set-Alias   -scope global -name   alias         -value   "Get-Alias"
Set-Alias   -scope global -name   help          -value   "Get-Help"
Set-Alias   -scope global -name   gethelp       -value   "Get-Help"
Set-Alias   -scope global -name   gindstr       -value   "findstr"
set-Alias   -scope global -name   srt           -value   "C:/Windows/System32/sort.exe"
set-Alias   -scope global -name   wsort         -value   "C:/Windows/System32/sort.exe"
Set-Alias   -scope global -name   usort         -value   "C:/Git/usr/bin/sort.exe"
Set-Alias   -scope global -name   o2s           -value   "ostr"
Set-Alias   -scope global -name   search        -value   "Select-String"
Set-Alias   -scope global -name   sea           -value   "Select-String"
Set-Alias   -scope global -name   grepp         -value   "Select-String"
Set-Alias   -scope global -name   ugrep         -value   "C:/Git/usr/bin/grep.exe"
Set-Alias   -scope global -name   rtdsql        -value   "rdtsql"
Set-Alias   -scope global -name   vf            -value   "setdir"
set-Alias   -scope global -name   xs            -value   "setdir"
Set-Alias   -scope global -name   cdd           -value   "setdir"
Set-Alias   -scope global -name   cd            -value   "setdir"               -force  -Option Allscope
Set-Alias   -scope global -name   diff          -value   "diff.exe"             -force  -Option Allscope
set-Alias   -scope global -name   wz            -value   "winzip32.exe"
Set-Alias   -scope global -name   so            -value   "Select-object"
Set-Alias   -scope global -name   cp            -value   "xcp"                  -force  -Option Allscope
set-Alias   -scope global -name   rm            -value   "xrm"                  -force  -Option Allscope
Set-Alias   -scope global -name   mv            -value   "xmv"                  -force  -Option Allscope
Set-Alias   -scope global -name   gc            -value   "xgc"                  -force  -Option Allscope
Set-Alias   -scope global -name   zz            -value   "xzz"                  -force  -Option Allscope
Set-Alias   -scope global -name   ?             -value   "where.exe"            -force  -Option Allscope
Set-Alias   -scope global -name   rr            -value   "InvokeDiaryCommand"   -force -Option Allscope
Set-Alias   -scope global -name   r             -value   "InvokeDiaryCommand"   -force -Option Allscope
Set-Alias   -scope global -name   h             -value   "__displayDiary"       -force -Option Allscope
Set-Alias   -scope global -name   xxr           -value   "ForbidDiaryCommand"   -force -Option Allscope
Set-Alias   -scope global -name   ij            -value   "idea64.exe"           -force -Option Allscope
Set-Alias   -scope global -name   idea          -value   "idea64.exe"           -force -Option Allscope
Remove-Item alias:ls
#
# }}}
# =========================================================================
#
# =========================================================================
# \\$domainUsrDir\$uid\powershell\helper_func.psml {{{
# =========================================================================
#

function ql { #{{{
     #
     # see section 7.1.2 (page 239)
     #
     #   of...
     #
     #   windows Powershell in Action, second Edition
     #   author BRUCE PAYETTE
     #   ISBN 9781935182139
     #
     # ...for a description of "function ql"
     $args
 } # }}}

function qlt { # {{{
    foreach($i in $(ql $args)) {
        echo $i
    }
} # }}}

function qs { #{{{
    #
    #  see section 7.1.2 (page 239)
    #
    #  of...
    #
    #  windows PowerSheil in Action, Second Edition
    #  author bruce payette
    #  ISBN 9781935182139
    #
    #  ...for a description of "function qs"
    "$args"
} # }}}

function buildKey { #{{{
      #
      #  this allows for a theoretical   maximum of
      #  10000 files loaded into a single vim session
      #
      "X" + "{0:D4}" -f $args[ 0 ]
} # }}}

function posh { #{{{
     #
     #  Command Line for windows
     #
     start-process $sys32/windowsPowerShell/v1.0/powershell.exe
} # }}}

function drives { # {{{
    Get-PSDrive -PSProvider FileSystem
} # }}}

function ShowErrorDetails { # {{{
    # see section 14.1.1 (page 559)
    #
    #   of...
    #
    #   windows Powershell in Action, second Edition
    #   author BRUCE PAYETTE
    #   ISBN 9781935182139
    #
    # ...for a description of "function ShowErrorDetails"
    param(
        $ErrorRecord = $Error[0]
    )

    $ErrorRecord | Format-List * -Force
    $ErrorRecord.InvocationInfo | Format-List *
    $Exception = $ErrorRecord.Exception
    for ( $depth = 0; $Exception -ne $null; $depth++ )
    {
        "$depth" * 80
        $Exception | Format-List -Force *
        $Exception = $Exception.InnerException
    }
} # }}}

function initBigRegexGBLS { # {{{

      # build regular expresion only once for larger string expressions
      #
      $pfxRe = '(?ix: Microsoft.PowerShell.Core\\FileSystem:: )?'

      $fullRE  = @()
      $fullRE += '(?x:   \A \s* Path \s* [-]+   \s* '     + $pfxRe + ' (?x: (.+?) \s* \z) )'
      $fullRE += '(?x:   \A \s* '                         + $pfxRe + ' (?x: (.+?) \s* \z) )'
      $fullRE += '(?ix:   \A (?x: System [.] )? String \s+   PSPath= ' + $pfxRe + ' (?x: (.+?) \s* \z) )'
      $fullRE += '(?six: \A ( (Env) | (variable) | (HKLM) | (HKCU) ) : )'

      $GLOBAL:getLocationRE           =  [regex]   $fullRE[   0  ]
      $GLOBAL:getPSPathRE             =  [regex]   $fullRE[   1  ]
      $GLOBAL:getLongNameRE           =  [regex]   $fullRE[   2  ]
      $GLOBAL:notAfileSystemRE        =  [regex]   $fullRE[   3  ]

 } # }}}

function currentDirectory { # {{{
      # $profile -replace '[\\][^\\]+\z', $nullstr
      #
      # will produce the name of the directory that contains...
      #
      #     Microsoft.PowerShell_profile.ps1
      #
      # if...
      #
      # $d == $profile -replace '[\\][^\\]+\z', $nullstr
      #
      # ...after cd $d...
      #
      # in some environments, Get-Location will
      # produce an object that is displayed
      # as follows:
      #
      # Path
      # ----
      # Microsoft.Powershell.Core\FileSystem::
      # \\$domainUsrDir\nomad\My Documents\WindowsPowerShell
      #
      # ...which is useless at best, or just plain *WRONG*
      #
      # A string (not an object) that displays as...
      #
      # \\$domainUsrDir\nomad\My Documents\WindowsPowershell
      #
      # ...is a lot more useful. That is what the code below does.
      #
      $( Get-Location | out-string ) -replace $getLocationRE, '$1'
 } # }}}

function ff1 { # {{{
    $GLOBAL:consoleTitleMode = "tskMgrInstance"
} # }}}

function setConsolewindowTitle { # {{{
    #
    #  display the current directory with
    #  embedded spaces and parens aliased
    #  to utf8 placibos
    #
    if ( "rt1path" -eq $consoleTitleMode ) {
        $host.UI.RawUI.WindowTitle = "rt1 ENVIRONMENT"
    }
    elseif ( "rt1devpath" -eq $consoleTitleMode ) {
        $host.UI.RawUI.windowTitle = "rt1 dev environment"
    }
    elseif ( "vsenvpath" -eq $consoleTitleMode ) {
        $host.UI.RawUI.windowTitle = "vs dev environment"
    }
    elseif ( "ClEnvPath" -eq $consoleTitleMode ) {
        $host.UI.RawUI.windowTitle = "CYTHON"
    }
    elseif ( "tskMgrInstance" -eq $consoleTitleMode ) {
        $host.UI.RawUI.windowTitle = "ADM " + $( currentDirDsply )
    }
    else {
        $host.UI.RawUI.WindowTitle = dspc $( currentDirectory )
    }
} # }}}

function currentDirDsply { # {{{
    #
    #  display the current directory with
    #  embedded spaces and parens aliased
    #  to utf8 placibos
    #
    dspc $( currentDirectory )
} # }}}

function rndx16bit { # {{{
      #
      #  return the random number
      #  zeropadded to 4 hex   digits
      #
      $("{0:x4}" -f $( Get-Random -maximum 0xFFFF )).ToUpper()
} # }}}

function rndx32bit { # {{{
      #
      # return   the   random number
      # zeropadded     to 8 hex digits
      #
      "{0:X8}" -f $( Get-Random )
} # }}}

function rnd32bit { # {{{
    #
    #  return the   random number
    #  zeropadded   to 10 decimal digits
    #
    "{0:d10}" -f $( Get-Random )
} # }}}

function pidl { #{{{
    #
    #  return the   process ID of the current process
    #  zeropadded   to 5 digits
    #
    "{0:d5}" -f $PID
} # }}}

function colorcheck { # {{{
    $f1 = "{0,-12}{1}"  # format specifier 1
    $f2 = "{0,-12}"     # format specifier 2
    [Enum]::Getvalues( $host.UI.RawUI.ForegroundColor.GetType() ) |
    foreach {
        write-Host                       -NoNewLine  "[ "
        write-Host                       -NoNewLine  $( $f1 -f $_,": " )
        write-Host  -ForegroundColor  $_ -NoNewLine  $( $f2 -f $_      )
        write-Host                                   " ]"
    }

    [Enum]::Getvalues( $host.UI.RawUI.BackgroundColor.GetType() ) |
    foreach {
        write-Host                       -NoNewLine  "[ "
        write-Host                       -NoNewLine  $( $f1 -f $_,": " )
        write-Host  -BackgroundColor  $_ -NoNewLine  $( $f2 -f $_      )
        write-Host                                   " ]"
    }
} # }}}

function remote2maggi { # {{{
    $s1 = "$cReturn$newline"
    $s2 = "${s1}ipmo C:\doc\WindowsPowerShell\PowerShell4Import.psm1$s1"
    Write-Host -ForegroundColor Green "$s2"
    $s3 = 'etsn -computername magdalene -credential magdalene\nomad'
    iex $s3
} # }}}

function remote2desales { # {{{
    $s1 = "$cReturn$newline"
    $s2 = "${s1}ipmo C:\doc\WindowsPowerShell\PowerShell4Import.psm1$s1"
    Write-Host -ForegroundColor Green "$s2"
    $s3 = 'etsn -computername desales -credential desales\nomad'
    iex $s3
} # }}}

function remote2fjs { # {{{
    $s1 = "$cReturn$newline"
    $s2 = "${s1}ipmo C:\doc\WindowsPowerShell\PowerShell4Import.psm1$s1"
    Write-Host -ForegroundColor Green $s2
    $s3 = 'etsn -computername fultonjsheen -credential fultonjsheen\ksk'
    iex $s3
} # }}}

function adbshell { # {{{
    $image = s2bs( "C:/adb/adb.exe shell" )
    invoke-expression -command $image
} # }}}

function gitgui { # {{{
    $d = $doublequote
    $image = s2bs( "start git/wish.exe ${d}C:/Git/mingw64/libexec/git-core/git-gui${d}" )
    invoke-expression -command $image
} # }}}

function isee {  # {{{
    # C:\windows\System32\WindowsPowerShell\v1.0\powershell_ise.exe
    Powershell_ISE "$profile,$profilem"
} # }}}

function ise { # {{{
    Powershell_ISE         $args
} # }}}

function pp { # {{{
      #
      # print short names of path constituents
      #
      $pathArray = $path -split ";"
      for ( $ix = 0; $ix -lt $pathArray.length; $ix++ ) {
          $entry = dspc $pathArray[ $ix ]
          "{0:d2} {1}" -f $ix, $entry
      }
} # }}}

function ppl { # {{{
     #
     # print long names of path constituents
     #
     $pathArray = $pathcache[ "path" ]
     for ( $ix = 0; $ix -lt $pathArray.length; $ix++ ) {
          $entry = dSpc $pathArray[ $ix ]
          "{0:d3} {1}" -f $ix, $(dSpc $(elementlongname $entry))
     }
} # }}}

function mststampInit { # {{{
     # FACILITATE THE FOLLOWING TRANSFORMATION...
     # TRANSFORM:   2011-07-28t18:22:41.2847383-04:00
     #      INTO:   20110728.182241.2847383
     #
     if ( $null -eq $tStampRE ) {
          $datePattern           = '(?six:  (\d+)  [-] (\d+)  [-] (\d+) )'
          $timePattern           = '(?six:  (\d\d) [:] (\d\d) [:] (\d\d) [.] (\d+) )'
          $GLOBAL:tStampRE       = "(?six:  \A $datePattern [T] $timePattern .+ \z)";
          $GLOBAL:tStemplate0    = '$1$2$3.$4$5$6.$7';
          $GLOBAL:tStemplate1    = '$1$2$3.$4$5$6';
          $GLOBAL:tStemplate2    = '$1$2$3$4$5';
     }
}    # }}}

function mststampshort { # {{{
     #
     # TRANSFORM: 2011-07-28t18:22:41.2847383-04:00
     #      INTO: 20110728.182241
     #
     $( get-date -format o ) -replace $tStampRE, $tStemplate1
}   # }}}

function mststampshorter { # {{{
     #
     # TRANSFORM: 2011-07-28t18:22:41.2847383-04:00
     #      INTO: 201107281822
     #
     $( get-date -format o ) -replace $tStampRE, $tStemplate2
}   # }}}

function mststamp { # {{{
     #
     # TRANSFORM: 2011-07-28t18:22:41.2847383-04:00
     #      INTO: 20110728.182241.2847383
     #
     $( get-date -format o ) -replace $tStampRE, $tStemplate0
}   # }}}

function usrHost { # {{{
     $env:USERNAME      +
     "."                +
     $localhost
}   # }}}

function usrHostPid { # {{{
    $( usrHost )  +
    "."           +
    $( pidl )
} # }}}

function usrHostTStampPidRE { # {{{
    $( usrHost )    +
    '[.]'           +
    '\d+'           +
    '[.]'           +
    '\d+'           +
    '[.]'           +
    '\d+'           +
    '[.]'           +
    $( pidl )
}   # }}}

function UidTStampRnd32 { # {{{
    $env:USERNAME         +
    $( mststampshorter )  +
    "."                   +
    $( rndx32bit )
} # }}}

function usrHostTStampRnd { # {{{
    $( usrHost )   +
    "."            +
    $( mststamp )  +
    "."            +
    $( rndx32bit )
} # }}}

function usrHostTStampPidRnd { # {{{
    $( usrHost )   +
    "."            +
    $( mststamp )  +
    "."            +
    $( pidl )      +
    "."            +
    $( rndx16bit )
} # }}}

function usrHostTStampPid { # {{{
    $( usrHost )        +
    "."                 +
    $( mststampshort )  +
    "."                 +
    $( pidl )
} # }}}

function usrHostPidTStamp { # {{{
    $( usrHostPid ) +
    "."             +
    $( mststamp )
} # }}}

function ostr { # {{{
    #
    #   transform the object passed in
    #   from a pipe, into a string
    #
    while ( $input.movenext() ) {
       $input.Current | Out-String -Stream
    }
}  # }}}

function bs2s { # {{{
    #
    #   replace dos-backslash with unix-frontslash
    #
    if ( $args.count ) {
        $target = qs $( [string]$args )
    }
    else {
        #
        #  No arguments given: Assume data was piped in.
        #
        $target = $( $input | Out-String )
    }
    $target -replace "$backslash$backslash", $slash
}  # }}}

function s2bs { # {{{
    #
    #   replace unix-frontslash with dos-backslash
    #
    if ( $args.count ) {
        $target = qs $( [string]$args )
    }
    else {
        #
        #  No arguments given: Assume data was piped in.
        #
        $target = $( $input | Out-String )
    }
    $target -replace $slash, $backslash
}  # }}}

function deSpaceEZ { # {{{
      #
      # Replace spaces and parens with unicode non-whitespace.
      # This is the simplest solution (EZ) but it is not as
      # general as the one I am using in the mainline code.
      #
      if ( 0 -ne $args.count )     {
           $s = qs $( [string]$args )
      }
      else   {
           $s = $( $input | Out-String )
      }
      $s = $s -replace '(?six:   \u0020  )',  $uSpaceAKA
      $s = $s -replace '(?six:   \u0028  )',  [string][char][int] 0x2039
      $s = $s -replace '(?six:   \u0029  )',  [string][char][int] 0x203A
      $s
} # }}}

function dspc { # {{{
    $input | deSpace @args
} # }}}

function enSpaceEZ { # {{{
    #
    #   REVERSE  THE DESPACE OPERATION
    #
    #   Replace unicode non-whitespace with their origina1 values.
    #   this is the simplest solution (EZ) but it is not as
    #   general as the one I am using in the mainline code.
    #
    if ( 0 -ne $args.count ) {
        $s =   qs $( [string]$args )
    }
    else {
        #
        # NO   ARGUMENT GIVEN: Assume    data  was piped in
        #
        $s =   $( $input | Out-String    )
    }
    $s   = $s -replace '(?six:     \u2022 )', [string][char][int] 0x0020
    $s   = $s -replace '(?six:     \u2039 )', [string][char][int] 0x0028
    $s   = $s -replace '(?six:     \u203A )', [string][char][int] 0x0029
    $s
} # }}}

function deSpace { # {{{
    #
    # SUBSTITUTE UNICODE SYMBOLS FOR WHITESPACE|PARENS IN PATHNAMES
    #
    if ( $args.count ) {
        $canonicalName = qs $( [string]$args )
    }
    else {
        #
        #  No arguments given: Assume data was piped in.
        #
        $canonicalName = $( $input | Out-String )
    }

    $garnishedName = $canonicalName
    if ( $garnishedName -match $regx2aka ) {
        foreach ( $xformRow in $chrAkaXform ) {
            #
            # Must accommodate posh braindamage below:
            # without '$( qs )' the replacement fails.
            #
            $gn = $garnishedName
            $gn = $gn -replace $xformRow[ 0 ], $( qs $xformRow[ 1 ] )
            $garnishedName = $gn
        }
    }

    bs2s $garnishedName # re-orient the backslash characters
} # }}}

function enspace { # {{{
    #
    # RESTORE PATHNAME TO ITZ ORIGINAL FORM
    # ( reverse the deSpace operation )
    #
    if ( $args.count ) {
        $garnishedName = qs $( [string]$args )
    }
    else {
        #
        # No argument given. Assume data was piped in.
        #
        $garnishedName = $( $input | Out-String )
    }

    $canonicalName = $garnishedName
    if ( $canonicalName -match $regx2org ) {
        foreach ( $xformRow in $chrAkaXform ) {
            #
            # Accomodate posh braindamage below:
            # without '$( qs )' the replacement fails.
            #
            $cn = $canonicalName
            $cn = $cn -replace $xformRow[ 2 ], $( qs $xformRow[ 3 ] )
            $canonicalName = $cn
        }
    }
    s2bs $canonicalName # re-orient the slash characters
} # }}}

function enspaceArgs { # {{{
    #
    # provide a way to automagically strip
    # the spaceAKA info from the input arguments
    # this is necessary for calls to:
    #
    # Get-Content ( gc ) *
    # Copy-Item   ( cp )
    # Move-item   ( mv )
    # Remove-Item ( rm )
    #
    # because those commands can take a comma separated
    # list of entries which posh treats in a special way.
    # This is currently (20120605) coded to handle only 1 level of
    # array nesting. I am not sure if any command uses
    # more than that. ( I cant even think of how it might look
    # on the command line.)
    #
    # * currently (20121210) 'gc -wait <file>'
    # fails with
    # Get-Content : A positional parameter cannot be found that accepts argument '.\basepath.bat'.
    # At \\$domainUsrDir\nomad\powershell\pshellMain.psml:683 char:13
    # + Get-Content <<<< @argsPrime
    # + Categorylnfo : invalidArgument: (:) [Get-Content], ParameterBindingException
    # + FullyQualifiedErrorid : PositionalParameterNotFound,Microsoft.PowerShell.Commands.GetContentcommand
    #
    #  Im not quite sure how to fix this.
    #
    #
    $argsPrime = @()
    $inpArgs = $args
    foreach ( $args1 in $inpArgs ) {
        if ( $args1.gettype().fullname -match 'System.Object\[\]' ) {
            #
            #   since args1 is itself a array nested in the args array,
            #   special  processing is necessary here
            #
            $args1Prime = @()
            foreach ( $args2 in $args1 ) {
                $args1Prime += enspace $args2
            }
            #
            #   without  the leading comma (below) we just append scalars
            #   the comma ensures that argsPrime gets another
            #   *ARRAY* appended to it.
            #
            $argsPrime += ,$args1Prime
        }
        else {
            $argsPrime += enspace $args1
        }
    }
    #
    # Among the many peculiar -- not  to say braindamaged --
    # conventions in posh is the way that it strips array
    # levels unless you precede $argsPrime with a comma
    #
    ,$argsPrime
} # }}}

function showAKAmatrix { # {{{
    foreach ( $xformRow in $chrAkaXform ) {
        $outStr = $nullstr
        foreach ( $entry in $xformRow ) {
            $outStr += "<< " + $entry + " >> "
        }
        Write-Host -ForegroundColor Green $outStr
    }

} # }}}

function nSpc { # {{{
    $input | enSpace @args
}    # }}}

function dblquote { # {{{
    $doublequote + $args + $doublequote
}    # }}}

function sqlquote { # {{{
    $singlequote + $args + $singlequote
}    # }}}

function xgc { # {{{
    #
    # an alias of 'gc' gets me here
    #
    $argsPrime = enspaceArgs @args
    Get-Content @argsPrime
} # }}}

function xrm { # {{{
    #
    # an alias of rm gets me here
    #
    $argsPrime = enspaceArgs @args
    remove-item @argsPrime
} # }}}

function xmv { # {{{
    #
    # an alias of mv gets me here
    #
    $argsPrime = enspaceArgs @args
    move-item @argsPrime
} # }}}

function xcp { # {{{
    #
    # an alias of cp gets me here
    #
    $argsPrime = enspaceArgs @args
    copy-item @argsPrime
} # }}}

function xzz { # {{{
    #
    # an alias of zz gets me here
    #
    $argsPrime = enspaceArgs @args
    7z.exe @argsPrime
} # }}}

function createTmpFolder { # {{{
    $ENV:temp = $ENV:userprofile + $backslash + "temp.$( mststamp )"
    $ENV:temp = $( s2bs $ENV:temp )
    New-Item -itemtype directory -path $ENV:temp
    Set-ItemProperty -Path HKCU:\Environment -Name temp -Value $ENV:temp
} # }}}

function guaranteeTmpDirExistence { # {{{
    #
    # if temp directory does not exist then
    # create it on the file system and create
    # a reference to it in the registry
    #
    if ( Test-Path -path ENV:/temp ) {
        if ( -not ( Test-Path -path $ENV:temp -PathType container ) ) {
            createTmpFolder
        }
    }
    else {
        createTmpFolder
    }
    $GLOBAL:temp = $( bs2s $ENV:temp )
} # }}}

function ffx { # {{{
    #
    # Run firefox
    #
    $argsPrime = enspaceArgs @args
    $argstring = nSpc $( qs $( [string]$argsPrime ) )
    if ( $nullstr -ne $argstring ) {
        if ( Test-Path -path $argstring -PathType leaf ) {
            $reader = shortname "$bin/MozFirefox/firefox.exe"
            . $reader  @argsPrime
        }
        else {
            $msg = "OUCH:<read>[ $argstring ] "
            $msg += "DOES NOT APPEAR TO BE A FILE."
            Write-Host -ForegroundColor Red $msg
        }
    }
    else {
        $reader = shortname "$bin/MozFirefox/firefox.exe"
        . $reader
    }
} # }}}

function read { # {{{
    #
    # Run adobe acrobat
    #
    $argsPrime = enspaceArgs @args
    $argstring = nSpc $( qs $( [string]$argsPrime ) )
    if ( $nullstr -ne $argstring ) {
        if ( Test-Path -path $argstring -PathType leaf ) {
            $acro11 = "$bin/Adobe/Acrobat11.0/Acrobat/Acrobat.exe"
            $reader = shortname "$bin/NitroReader3/NitroPDFReader.exe"
            if ( Test-Path -path $acro11 -PathType leaf ) {
                $reader = shortname $acro11
            }
            . $reader $(Get-Item $argstring)
        }
        else {
            $msg = "OUCH:<read>[ $argstring ] "
            $msg += "DOES NOT APPEAR TO BE A FILE."
            Write-Host -ForegroundColor Red $msg
        }
    }
    else {
        $msg = $newline + 'OUCH:<read>[ $argstring ] '
        $msg += 'IS EMPTY. READ WHAT?'
        write-Host -ForegroundColor Red $msg
    }
} # }}}

function elementlongname { # {{{
     #
     # return 8.3 filename for a folder
     #
     $target = nSpc $( s2bs $( qs $( [string]$args ) ) )

     $defstring = $( Get-Item $target                 |
                     Get-Member -name "PSPath"        |
                     %{ $_.definition }               )

     bs2s $( $defstring -replace $getLongNameRE, '$1' )
} # }}}

function fileshortname { # {{{
    #
    # return 8.3 filename for   a file
    #
    $target = nSpc $( s2bs $( qs $( [string]$args ) ) )
    $fileobj = $fso.GetFile( $target )
    $fileobj.ShortPath
} # }}}

function foldershortname { # {{{
    #
    # return 8.3 filename for  a folder
    #
    $target = nspc $( s2bs $( qs $( [string]$args ) ) )
    $folderobj = $fso.GetFolder( $target )
    $folderobj.ShortPath
} # }}}

function shortname { # {{{
    #
    # return 8.3 filename for a folder or file if it exists
    #
    $target = nspc $( qs $( [string]$args ) )
    if ( Test-Path -path $target -PathType leaf ) {
        fileshortname $target
    }
    elseif ( Test-path -path $target -PathType container ) {
        foldershortname $target
    }
    else {
        $msg = "OUCH<shortname>:[ $target        ] "
        $msg += "DOES NOT APPEAR TO BE A FILE OR DIRECTORY"
        write-Host -Foregroundcolor Red $msg
        write-Output $nullstr
        $nullstr
    }
} # }}}

function longname { # {{{
    #
    # return 8.3 filename for a folder or file if it exists
    #
    $target = nSpc $( qs $( [string]$args ) )
    if ( Test-Path -path $target ) {
        elementlongname $target
    }
    else {
        $msg = "OUCH<longname>:[ $target ] "
        $msg += "DOES NOT APPEAR TO BE A FILE OR DIRECTORY"
        write-Host -Foregroundcolor Red $msg
        write-Output $nullstr
        $nullstr
    }
} # }}}

function touch { # {{{
    #
    # change modification date for a file
    # to the current date and time. If
    # the file does not exist, then create
    # an empty one.
    #
    $timenow = Get-Date
    foreach ( $maybeAfile in $( ql $( $args ) ) ) {
        $target = nSpc $( s2bs $( qs $( $maybeAfile ) ) )
        if ( Test-Path -path $target -PathType leaf ) {
            $fileObject = Get-Item $target
            $fileobject.LastwriteTime = $timenow
        }
        else {
            Add-Content -path $target -value $nullstr
            Clear-Content -path $target
            if ( Test-Path -path $target -PathType leaf ) {
                $fileObject = Get-Item $target
                $fileobject.LastwriteTime = $timenow
            }
        }
    }
} # }}}

function shortnamequiet { # {{{
    #
    # return 8.3 filename for a folder or file
    # dont complain if it does not exist
    #
    $target = nSpc $( qs $( [string]$args ) )
    if ( Test-Path -path $target -PathType leaf ) {
        fileshortname $target
    }
    elseif ( Test-Path -path $target -PathType container ) {
        foldershortname $target
    }
    else {
        $target
    }
} # }}}

function cleave { # {{{
    $inpArg = $( qs $( [string]$args ) )

    [System.Collections.ArrayList]$SCA = $inpArg -split '[\n\r;]+'          |
                                         %{ $_ -replace $leadingWS, $nullstr }  |
                                         %{ $_ -replace $trailingWS, $nullstr }
    [System.Collections.ArrayList]$SCA
} # }}}

function sunder { # {{{
    $inpArg = $( qs $( [string]$args ) )

    $inpArg -split '[\n\r;]+'          |
    %{ $_ -replace $leadingWS, $nullstr }  |
    %{ $_ -replace $trailingWS, $nullstr }

} # }}}

function sow { # {{{
    #
    # join the members of the array (second argument)
    # into a semicolon separated string whose identifier variable name
    # (without the '$' prefix) is supplied in the string (first argument)
    # but only if the argument exists as a directory accessible
    # by the local host. The variable '<TAG>' gets the short named
    # path. The variable '<TAG>L' gets the long name path.
    #
    # Also store the '<TAG>L' array in $GLOBAL:pathCache[ $tag ]
    # and save the long and short string in a corresponding
    # environment variable.
    #
    # clean up the entries in the array and skip the ones that are
    # commented out ( start with # )
    #
    if ( 1 -gt $args.count ) {
        $m = @()
        $m+= $newline
        $m+= "OUCH<sow>:[ sow ]"
        $m+= $newline
        $m+= 'USAGE: [ sow tag value ]'
        $m+= $newline
        $errusage = $m -join $newline
        Write-Host -Foregroundcolor Red $errusage
        return
    }
    $tag = $args[0]                                 #  1st Argument
    $tag = $tag  -replace  $leadingWS,  $nullstr    #  trim  leading white space
    $tag = $tag  -replace  $trailingWS, $nullstr    #  trim  trailing white space
    #
    # $sowTypeAllowed is a global regex that only lets through
    # the desired types (Array|String) the introspection in
    # posh aint great but it is better than nothing.
    #
    if ( 2 -gt $args.count ) {
        $args = (, $args[0], "")
    }
    else {
        if ( !$args[1] ) {
            $args[1] = ""
        } 
    }
    
    $argType = $($args[1].psobject.typenames | where {$_ -match $sowTypeAllowed})
    if ( "System.String" -eq $argType ) {
        $pathArray = (, $args[1])
    }
    elseif ( $null -ne $argType ) {
        $pathArray = $args[1]
    }
    else {
        $argType = $args[1].psobject.typenames -join $space
        throw "Invalid Argument ( $argType ) supplied to sow"
    }
    if ( $nullstr -eq $args[1] ) {
        $m = "skipping sow for " + $args[0]
        Write-Host -Foregroundcolor Red $m
        return
    }
    $pVfydSFN = @()                                 #  path ( existence verified ) short folder name
    $pVfydLFN = @()                                 #  path ( existence verified ) long  folder name
    $uniqFolders = @{}                              #  keep duplicate entries out of path
    foreach ( $entry in $pathArray ) { # {{{
        $entry    =  $entry -replace  '[#].*',      $nullstr    #  trim  pound sign and data after
        $entry    =  $entry -replace  $leadingWS,   $nullstr    #  trim  leading white space
        $entry    =  $entry -replace  $trailingWS,  $nullstr    #  trim  trailing white space

        if ( $entry -match '\A\s*\z' ) {
            continue      #  empty, so skip
        }

        if ( $entry -match '\A\s*#' ) {
            continue       #  commented out,   so  skip
        }

        if ( $entry -match '\A[/\\][^/\\]' ) {
            #  Leading slash is artifact of nonexistant directory.
            # Skip it unless there are 2 of them.
            continue
        }

        $entry = bs2s $entry # normalize path separator

        if ( "tracethis" -eq "tracethisx" ) {
            $m = "path redundancy test:" + "[$entry]"
            Write-Host -Foregroundcolor Green $m
        }
        if ( -not ( Test-Path -path $entry -PathType container ) ) {
            continue       #  does not exist as a directory on this machine
        }

        if ( $null -ne $uniqFolders[ $entry ] ) {
            continue       #  already in the path
        }
        #
        # mark this folder so it does not show up
        # in the path in more than 1 place
        #
        $uniqFolders[ $entry ] = 1
        #
        # push the long folder name on the array
        # since itz existence has been verified
        #
        $pVfydLFN += $entry
        #
        # push the short folder name on the array
        #
        $pVfydSFN += bs2s $( foldershortname  $entry )

    } # }}}

    if ( 0 -ne $pVfydLFN.length ) {
        #
        # only set value if path has something in it
        #
        $GLOBAL:pathcache[ $tag ] = $pVfydLFN
        $orderedPair  = ,$( $pVfydSFN -join ';' )
        $orderedPair +=  $( $pVfydLFN -join ';' )
        foreach ( $prefix in $( ql global env ) ) {
            $xpression = '${0}:{1}, ${0}:{1}L = ${2}' -f $(ql $prefix $tag orderedPair)
            #
            # Say What?
            #
            # invoked expression will look like one
            # of the lines below. Thereby assuring
            # that each WINDOWS ENV: variable has
            # a corresponding powershell GLOBAL variable.
            # ordered pair contains 2 path strings the first
            # one is harder for most humans to read than the
            # second.
            #
            # $GLOBAL:INCLUDE, $GLOBAL:INCLUDEL = $orderedPair
            #    $ENV:INCLUDE,    $ENV:INCLUDEL = $orderedPair
            #
            invoke-Expression $xpression
        }
    }

} # }}}

function ingest { # {{{
    # take a single argument of strings separated
    # by linefeeds. The first string is a pathName
    # like "INCLUDE" the remaining strings are
    # directory paths. each directory will be added
    # to search path for pathName. String will look like:
    #
    # INCLUDE
    # $bin/VisualStudio14_0/VC/INCLUDE
    # $pfdL/Windows Kits/10/include/10.0.10240.0/ucrt
    # $pfdL/Windows Kits/NETFXSDK/4.6/include/um
    # $pfdL/Windows Kits/10/include/10.0.10240.0/shared
    # $pfdL/Windows Kits/10/include/10.0.10240.0/um
    # $pfdL/Windows Kits/10/include/10.0.10240.0/winrt#
    #
    # ...But without the '#' character
    #
    [System.Collections.ArrayList]$pathGroup = cleave $args[0]
    $pName = $pathGroup[0]
    $pathGroup.RemoveAt(0)
    sow $pName $pathGroup
} # }}}

function ePlant { # {{{
    #
    # Set a global variable and an environment variable
    # to a value supplied in the second argument.
    #
    # This helps to keep the dos environment and the powershell
    # environments in sync.
    #
    # The first argument contains a string (usually without the `$` sign in front)
    # that will be the env and global variable name.
    #
    # Example:
    #
    #     ePlant consoleTitleMode 'ClEnvPath'
    #
    # Will set... $GLOBAL:consoleTitleMode = 'ClEnvPath'
    #                $ENV:consoleTitleMode = 'ClEnvPath'
    #
    $tag    = $args[0]                                     #  1st Argument
    $tag    = $tag    -replace  $leadingWS,    $nullstr    #  trim  leading white space
    $tag    = $tag    -replace  $trailingWS,   $nullstr    #  trim  trailing white space
    $value  = $args[1]                                     #  2nd Argument
    $value  = $value  -replace  $leadingWS,    $nullstr    #  trim  leading white space
    $value  = $value  -replace  $trailingWS,   $nullstr    #  trim  trailing white space

    foreach ( $prefix in $( ql global env ) ) {
        $xpression = '${0}:{1} = ${2}' -f $(ql $prefix $tag value)
        #
        # Say What?
        #
        # invoked expression will look like one
        # of the lines below. Thereby assuring
        # that each WINDOWS ENV: variable has
        # a corresponding powershell GLOBAL variable.
        #
        # $GLOBAL:TAG = $value
        #    $ENV:TAG = $value
        #
        invoke-Expression $xpression
    }
} # }}}

function ttff { # {{{
        $prefix = UidTStampRnd32
        bs2s "$sundryTmpD/$prefix.cointoss"
} # }}}

function sha1 { # {{{
    #
    # generate SHA1 sum
    #
    $caught = 0
    $m = @()
    $m+= $newline
    $m+= 'OUCH<sha1>:[ sha1 $argstring ]'
    $m+= $newline
    $m+= "USAGE: sha1 [ <filename> ] "
    $m+= $newline
    $errusage = $m -join $newline

    try {
        $argstring = nSpc $( bs2s $( qs $( [string]$args ) ) )

        $prefix = usrHostTStampPid
        $file2check = bs2s "$sundryTmpD/$prefix.cointoss"

        if ( $nullstr -eq "$argstring" ) {
            throw
            #
            #   head called without arguments ( In prior version )
            #
            $input | Out-String | Set-Content $file2check
        }
        elseif ( Test-Path -path $argstring -PathType leaf ) {
            #
            # the entire argument string is a filename
            #
            $file2check = $argstring
        }
        else {
            #
            # $argstring is *NOT* a file on the file system
            #
            $m    =  $nullstr
            $m   +=  $newline
            $m   +=  "OUCH<sha1>:[ $argstring  ] "
            $m   +=  "DOES NOT APPEAR  TO  BE A  FILE."
            $m   +=  $newline
            Write-Host -Foregroundcolor Red $m
            throw
        }
    }
    catch   {
        write-Host -Foregroundcolor Red $errusage
        $caught   =  1
    }


    if ( 0 -eq $caught ) {
        ## create the hash object that calculates the sha1
        $hashMethod = [Type] "System.Security.Cryptography.SHA1"
        $engine = $hashMethod::Create()

        $fObj = get-item $file2check
        $datastream = New-object io.streamReader $fobj.fullname
        $octets = $engine.ComputeHash( $dataStream.BaseStream )
        $dataStream.Close()
        $transform = New-object System.Text.StringBuilder
        $octets | Foreach-object {
            [void] $transform.Append( $_.ToString("x2") )
        }
        $transform.ToString()
    }

} # }}}

function s256 { # {{{
    #
    # generate sha256 sum
    #
    $caught = 0
    $m = @()
    $m+= $newline
    $m+= 'OUCH<s256>:[ s256 $argstring ]'
    $m+= $newline
    $m+= "USAGE: s256 [ <filename> ] "
    $m+= $newline
    $errusage = $m -join $newline

    try {
        $argstring = nSpc $( bs2s $( qs $( [string]$args ) ) )

        $prefix = usrHostTStampPid
        $file2check = bs2s "$sundryTmpD/$prefix.cointoss"

        if ( $nullstr -eq "$argstring" ) {
            throw
            #
            #   head called without arguments ( In prior version )
            #
            $input | Out-String | Set-Content $file2check
        }
        elseif ( Test-Path -path $argstring -PathType leaf ) {
            #
            # the entire argument string is a filename
            #
            $file2check = $argstring
        }
        else {
            #
            # $argstring is *NOT* a file on the file system
            #
            $m    =  $nullstr
            $m   +=  $newline
            $m   +=  "OUCH<s256>:[ $argstring  ] "
            $m   +=  "DOES NOT APPEAR  TO  BE A  FILE."
            $m   +=  $newline
            Write-Host -Foregroundcolor Red $m
            throw
        }
    }
    catch   {
        write-Host -Foregroundcolor Red $errusage
        $caught   =  1
    }


    if ( 0 -eq $caught ) {
        ## create the hash object that calculates the md5
        $hashMethod = [Type] "System.Security.Cryptography.SHA256"
        $engine = $hashMethod::Create()

        $fObj = get-item $file2check
        $datastream = New-object io.streamReader $fobj.fullname
        $octets = $engine.ComputeHash( $dataStream.BaseStream )
        $dataStream.Close()
        $transform = New-object System.Text.StringBuilder
        $octets | Foreach-object {
            [void] $transform.Append( $_.ToString("x2") )
        }
        $transform.ToString()
    }

} # }}}

function s512 { # {{{
    #
    # generate sha512 sum
    #
    $caught = 0
    $m = @()
    $m+= $newline
    $m+= 'OUCH<s512>:[ s512 $argstring ]'
    $m+= $newline
    $m+= "USAGE: s512 [ <filename> ] "
    $m+= $newline
    $errusage = $m -join $newline

    try {
        $argstring = nSpc $( bs2s $( qs $( [string]$args ) ) )

        $prefix = usrHostTStampPid
        $file2check = bs2s "$sundryTmpD/$prefix.cointoss"

        if ( $nullstr -eq "$argstring" ) {
            throw
        }
        elseif ( Test-Path -path $argstring -PathType leaf ) {
            #
            # the entire argument string is a filename
            #
            $file2check = $argstring
        }
        else {
            #
            # $argstring is *NOT* a file on the file system
            #
            $m    =  $nullstr
            $m   +=  $newline
            $m   +=  "OUCH<s512>:[ $argstring  ] "
            $m   +=  "DOES NOT APPEAR  TO  BE A  FILE."
            $m   +=  $newline
            Write-Host -Foregroundcolor Red $m
            throw
        }
    }
    catch   {
        write-Host -Foregroundcolor Red $errusage
        $caught   =  1
    }


    if ( 0 -eq $caught ) {
        ## create the hash object that calculates the checksum
        $hashMethod = [Type] "System.Security.Cryptography.SHA512"
        $engine = $hashMethod::Create()

        $fObj = get-item $file2check
        $datastream = New-object io.streamReader $fobj.fullname
        $octets = $engine.ComputeHash( $dataStream.BaseStream )
        $dataStream.Close()
        $transform = New-object System.Text.StringBuilder
        $octets | Foreach-object {
            [void] $transform.Append( $_.ToString("x2") )
        }
        $transform.ToString()
    }

} # }}}


function m5 { # {{{
    #
    # generate md5 sum
    #
    $caught = 0
    $m = @()
    $m+= $newline
    $m+= 'OUCH<m5>:[ m5 $argstring ]'
    $m+= $newline
    $m+= "USAGE: m5 [ <filename> ] "
    $m+= $newline
    $m+= "* IF <filename> ARGUMENT IS OMITTED, STDIN IS ASSUMED."
    $m+= $newline
    $m+= "* FILE NAMES CONTAINING SPACES ARE NOT SUPPORTED."
    $m+= $newline
    $errusage = $m -join $newline

    try {
        $argstring = nSpc $( bs2s $( qs $( [string]$args ) ) )

        $prefix = usrHostTStampPid
        $file2check = bs2s "$sundryTmpD/$prefix.cointoss"

        if ( $nullstr -eq "$argstring" ) {
            #
            #   head called without arguments
            #
            $input | Out-String | Set-Content $file2check
        }
        elseif ( Test-Path -path $argstring -PathType leaf ) {
            #
            # the entire argument string is a filename
            #
            $file2check = $argstring
        }
        else {
            #
            # $argstring is *NOT* a file on the file system
            #
            $m    =  $nullstr
            $m   +=  $newline
            $m   +=  "OUCH<m5>:[ $argstring  ] "
            $m   +=  "DOES NOT APPEAR  TO  BE A  FILE."
            $m   +=  $newline
            Write-Host -Foregroundcolor Red $m
            throw
        }
    }
    catch   {
        write-Host -Foregroundcolor Red $errusage
        $caught   =  1
    }


    if ( 0 -eq $caught ) {
        ## create the hash object that calculates the md5
        $hashMethod = [Type] "System.Security.Cryptography.MD5"
        $engine = $hashMethod::Create()

        $fObj = get-item $file2check
        $datastream = New-object io.streamReader $fobj.fullname
        $octets = $engine.ComputeHash( $dataStream.BaseStream )
        $dataStream.Close()
        $transform = New-object System.Text.StringBuilder
        $octets | Foreach-object {
            [void] $transform.Append( $_.ToString("x2") )
        }
        $transform.ToString()
    }

} # }}}

function b64df { # {{{
    $errmsg = $null
    :out do {
        if ( ( 1 -ne $args.count ) -and ( 2 -ne $args.count ) ) {
            $errmsg =  "I NEED 1 ARGUMENT OR 2"
            break out
        }
        if ( 1 -eq $args.count ) {
            $file2decode = $args[ 0 ]
            $expression = [regex] '(?six: [.] b64 \z )'
            $name4decodedfile = $file2decode -replace  $expression, ""
            if ( $file2decode -eq $name4decodedfile ) {
                $errmsg = "I NEED A DESTINATION FILE NAME"
                break out
            }
        }
        else {
            $file2decode = $args[ 0 ]
            $name4decodedfile = $args[ 1 ]
        }
        if ( -not ( Test-Path -path $file2decode -PathType leaf ) ) {
            $errmsg = "$file2decode DOES NOT SEEM TO BE A FILE"
            break out
        }
        if ( Test-Path -path $name4decodedfile -PathType leaf ) {
            $prefix = "$(rndx32bit)."
            $backupfile = "$prefix$name4decodedfile"
            move-item $name4decodedfile $backupfile
        }
        certutil -decode $file2decode $name4decodedfile
    } while ( $false )

    if ( $errmsg ) {
        $m = $nullstr
        $m += $newline
        $m += $errmsg
        $m += $newline
        write-Host -ForegroundColor Red $m
    }
} # }}}

function b64ef { # {{{
    $file2encode = $args[ 0 ]
    $ext = ".b64"
    $name4encodedfile = "$file2encode$ext"
    if ( Test-Path -path $name4encodedfile -PathType leaf ) {
        $prefix = "$(rndx32bit)."
        $backupfile = "$prefix$name4encodedfile"
        move-item $name4encodedfile $backupfile
    }
    certutil -encode $file2encode $name4encodedfile
} # }}}

function m5f { # {{{
    $inparg = $args
    $canonicalName = nSpc $( bs2s $( qs $( [string]$inparg ) ) )
    $r = m5 $canonicalName
    "$r $inparg"
} # }}}

function H_T_Proxy { # {{{
    try {
        $argstring = nSpc $( bs2s $( qs $( [string]$args ) ) )

        $commandStack = @()
        #
        # the trick below is the only way I found to extract
        # all the meaningful data from the call stack without
        # dumping it out to a file and re-parsing it.
        #
        $oCallStack = Get-PSCallStack | Format-List | Out-String
        # $oCallStack
        #
        # after "$oCallstack = Get-PSCallstack | Format-List | Out-string"
        # oCallStack will contain something like the data   below...
        #
        #
        #   scriptName         :  \\$domainUsrDir\nomad\powershell\pShellMain.psml
        #   ScriptLineNumber   :  321
        #   InvocationInfo     :  System.Management.Automation.lnvocationlnfo
        #   Command            :  H_T_Proxy
        #   Location           :  pshellMain.psml:  Line 321
        #   Arguments          :  {vanillaenv.txt}
        #
        #   ScriptName         :  \\$domainUsrDir\nomad\powershell\pShellMain.psml
        #   ScriptLineNumber   :  472
        #   InvocationInfo     :  System.Management.Automation.lnvocationlnfo
        #   Command            :  head
        #   Location           :  pShellMain.psml:  Line 472
        #   Arguments          :  {vanillaenv.txt}
        #
        #   scriptName         :
        #   ScriptLineNumber   :  1
        #   InvocationInfo     :  System.Management.Automation.InvocationInfo
        #   Command            :  prompt
        #   Location           :  prompt
        #   Arguments          :  {}
        #
        #
        $oCallStack -split '\n' | foreach {
            if (   $_ -match   '(?six: \A Command )' ) {
                #
                # find the caller id
                #
                $commandStack += $_
            }
        }
        $pattern = [regex]'(?six: \A \s* command \s+ [:] \s* ( ( head ) | ( tail ) ) \s* \z )'
        if ( $commandStack[ 1 ] -match $pattern ) {
            $callerlD = [string]$matches[ 1 ]
        }
        else {
            $m =   @()
            $m+=   $newline
            $m+=   "OUCH< H_T_Proxy>:[ H_T_Proxy $argstring ]"
            $m+=   $newline
            $m+=   " H_T_Proxy ONLY ACCEPTS CALLS FROM FUNCTIONS NAMED 'head' OR 'tail'"
            $m+=   $newline
            $errusage = $m -join $newline
            throw
        }


        $m = @()
        $m+= $newline
        $m+= "OUCH< H_T_Proxy>:[ head $argstring ]"
        $m+= $newline
        $m+= "USAGE: $callerlD [ [ [-]<linecount> ] [ <filename> ] ]"
        $m+= $newline
        $m+= "* IF <filename> ARGUMENT IS OMITTED, STDIN IS ASSUMED."
        $m+= $newline
        $m+= "* FILE NAMES CONTAINING SPACES ARE NOT SUPPORTED."
        $m+= $newline
        $errusage = $m -join $newline

        $caught = 0
        $linecount = 16
        $prefix = usrHostTStampPid
        $tmpfname = bs2s "$sundryTmpD/$prefix.cointoss"

        if ( $nullstr -eq $argstring ) {
            #
            # head called without arguments
            #
            $input | Out-String | Set-Content -Encoding UTF8 $tmpfname
        }
        elseif ( Test-Path -path $argstring -PathType leaf ) {
            #
            # the entire argument string is a filename
            #
            Get-content $argstring                        |
            Out-String                                    |
            Set-Content -Encoding UTF8 $tmpfname
        }
        elseif ( $argstring -notmatch '(?six: \A (?:[-]){0,2} (\d+) )' ) {
            #
            #  the  argument  string  does not begin with a linecount specifier
            #
            $m = $nullstr
            $m += $newline
            $m += "OUCH< H_T_Proxy>: [ $argstring ] "
            $m += "DOES NOT APPEAR TO BE A FILE."
            $m += $newline
            write-Host -ForegroundColor Red $m
            throw
        }
        elseif ( $argstring -match '(?six: \A (?:[-]){0,2}  (\d+) \z )' ) {
            #
            # the argument string consists of *only* a linecount specifier
            #
            $linecount = [int]$matches[ 1 ]
            $input | Out-String | Set-Content -Encoding UTF8 $tmpfname
        }
        elseif ( $argstring -match '(?six: \A (?:[-]){0,2} (\d+) \s+ (\S.*) \z )' ) {
            #
            # the argument string begins with a linecount specifier but has
            # data after the specifier
            #
            $linecount = [int]$matches[ 1 ]
            $secondarg = [string]$matches[ 2 ]
            if ( Test-Path -path $secondarg -PathType leaf ) {
                #
                # the data after the linecount specifier is a
                # file on the file system
                #
                Get-Content $secondarg                         |
                Out-String                                     |
                Set-Content -Encoding UTF8 $tmpfname

            }
            else {
                #
                # the data after the linecount specifier
                # is *NOT* a file on the file system
                #
                $m = $nullstr
                $m += $newline
                $m += "OUCH< H_T_Proxy>:[ $secondarg ] "
                $m += "DOES NOT APPEAR TO BE A FILE."
                $m += $newline
                write-Host -ForegroundColor Red $m
                throw
            }
        }
        else {
            #
            # the arguments somehow evaded all the above tests
            #
            throw
        }
    }
    catch   {
        write-Host -ForegroundColor Red  $errusage
        $caught = 1
    }

    if ( 0 -eq $caught ) {
        #
        # whether heads or tails is the  outcome,
        # filter-out/remove blank lines
        #
        if ( 'head' -eq $callerlD ) {
            Get-content $tmpfname                                    |
            Out-String -stream                                       |
            Select-String -notmatch '(?six: ^ \s* $ )'               |
            Select-Object -first $linecount
        }
        else {
            Get-Content $tmpfname                                    |
            Out-String -stream                                       |
            Select-String -notmatch '(?six: ^ \s* $ )'               |
            Select-Object -last $linecount
        }
    }

} # }}}

function head { # {{{
    $input | H_T_Proxy @args
} # }}}

function tail { # {{{
    $input |  H_T_Proxy @args
} # }}}
#
# }}}
# =========================================================================
#

function appendRecentEntries2Diary { # {{{
    $whereIwuz = get-location
    $epoch = 0
    $originalCount = $diary.count
    if ( Test-Path -path $psDiaryXmlFile -PathType leaf )
    { # {{{
        $referenceFile = Get-Item $psDiaryXmlFile
        $epoch = $referenceFile.LastWriteTime
    } # }}}
    set-location $psDiaryRecent
    Get-ChildItem *.pshcmd.xml  |
    sort-object LastWriteTime   |
    foreach {
        #
        # Gather commands and execution times into an ordinary array
        # because I cannot be sure 2 different commands do not have
        # the same timestamp.
        #
        $laststamp = 0
        if ( 0 -lt $diary.count ) {
            $lastIndex = $diary.count - 1
            $lastStamp =  $GLOBAL:diary[ $lastIndex ].stamp
            if ( $_.lastwritetime -gt $epoch )
            { # {{{
                $histobj = $null
                $GLOBAL:ErrorActionPreference = "SilentlyContinue"
                $GLOBAL:DebugPreference = "SilentlyContinue"
                $histobj = import-clixml $_
                $GLOBAL:ErrorActionPreference = "Inquire"
                $GLOBAL:DebugPreference = "Inquire"
                if ( $histobj -ne $null ) {
                    try {
                        $diaryRecord = [pscustomobject]@{
                                     stamp = $histobj.StartExecutionTime.ticks.tostring()
                               commandline = $histobj.commandline
                                    origin = $_.name
                        }
                        if ( $diaryRecord.stamp -gt $lastStamp )
                        { # {{{
                            $GLOBAL:diary += ,( $diaryRecord )
                        } # }}}
                    }
                    catch {
                        $msg = "OUCH:<appendRecentEntries2Diary>"
                        $msg += "$cReturn$newline"
                        $msg += "$cReturn$newline"
                        $msg += "[ $_.fullname ] "
                        $msg += "$cReturn$newline"
                        $msg += "$cReturn$newline"
                        $msg += "APPEARS TO HAVE A PROBLEM"
                        Write-Host
                        Write-Host -ForegroundColor Red $msg
                        Write-Host
                        llff $_.fullname
                    }
                }
                else {
                    $msg = "OUCH:<appendRecentEntries2Diary>"
                    $msg += "$cReturn$newline"
                    $msg += "$cReturn$newline"
                    $msg += "[ $_.fullname ] "
                    $msg += "$cReturn$newline"
                    $msg += "$cReturn$newline"
                    $msg += "APPEARS TO HAVE A PROBLEM"
                    Write-Host
                    Write-Host -ForegroundColor Red $msg
                    Write-Host
                    llff $_.fullname
                }
                $GLOBAL:ErrorActionPreference = "Inquire"   # Other setting == "SilentlyContinue"
                $GLOBAL:DebugPreference = "Inquire"         # Other setting == "SilentlyContinue"
            } # }}}
        }
    }
    if ( $originalCount -ne $diary.count ) {
        try {
            $savefailed = $false
            savediary
        }
        catch {
            $savefailed = $true
            $m = $nullstr
            $m += $newline
            $m += "*** SAVEDIARY FAILED ***"
            $m += $newline
            write-host -foregroundcolor magenta $m
        }
        finally {
            if ( -not $savefailed ) {
               Move-Item *.pshcmd.xml $psDiaryArchive
            }
        }
    }
    set-Location $whereIwuz
} # }}}

function loadDiary { # {{{
    $GLOBAL:diary = $null
    $GLOBAL:diary = Import-CliXml $psDiaryXmlFile
} # }}}

function saveDiary { # {{{
    $xmltmpfile = $psDiaryLclD + '/' + $(usrHostTStampPid) + ".xml"
    $xmltmpbase = $psDiaryLclD + '/' + $(usrHostTStampPid)
    Export-Clixml -Path $xmltmpfile -InputObject $diary
    Copy-Item $xmltmpfile $psDiaryXmlFile
    $xpression = "$bzipexe $xmltmpfile"
    $GLOBAL:ErrorActionPreference = "SilentlyContinue"
    $GLOBAL:DebugPreference = "SilentlyContinue"
    try {
        invoke-Expression $xpression
        $m = $nullstr
        $m += $newline
        $m += $(dspc $xmltmpfile)
        $m += ".bz2"
        $m += $newline
        write-host -foregroundcolor yellow $m
    }
    catch {
        write-host -foregroundcolor magenta "**ERROR: [ $xpression ] failed :RORRE**"
    }
    $GLOBAL:ErrorActionPreference = "Inquire"
    $GLOBAL:DebugPreference = "Inquire"
} # }}}

function dumpDiary { # {{{
    $txttmpfile = $psDiaryLclD + '/' + $(usrHostTStampPid) + ".txt"
    $stream = [System.IO.StreamWriter] $txttmpfile
    for( $i = 0 ; $i -lt $diary.count; $i++ )
    { # {{{
         $stream.WriteLine( ("H{0:D6}: {1}" -f $i, $diary[$i].commandline) )
    } # }}}
    $stream.close()
    Copy-Item $txttmpfile $psDiaryTxtFile
    $m = $nullstr
    $m += $newline
    $m += "( $i ) "
    $m += $(dspc $txttmpfile)
    $m += ".bz2"
    $m += $newline
    write-host -foregroundcolor yellow $m
} # }}}

function dumpDiarySlowly { # {{{
    $txttmpfile = $psDiaryLclD + '/' + $(usrHostTStampPid) + ".txt"
    $lineCount = 0
    $bigString = $nullStr
    foreach( $diaryRecord in $GLOBAL:diary )
    { # {{{
        $line = "H{0:D6}: {1}" -f $lineCount++, $diaryRecord.commandline
        Write-Output $line | Out-File -FilePath $txttmpfile -Append -Encoding UTF8
        if ( 0 -eq $lineCount % 512 )
        { # {{{
            write-Host -ForegroundColor Green "line count: $lineCount"
        } # }}}
    } # }}}
    $m = $nullstr
    $m += $newline
    $m += $(dspc $txttmpfile)
    $m += $newline
    write-host -foregroundcolor yellow $m
    Copy-Item $txttmpfile $psDiaryTxtFile
} # }}}

function saveSomeDiaryCommand { # {{{
    $w2s = "$psDiaryRecent/" + "$( usrHostTStampPid ).pshcmd.xml"
    $where2save = $w2s
    $target = qs $( [string]$args )
    Export-Clixml -Path $where2save -InputObject $diary[$target]
} # }}}

function generateChronicle { # {{{
    #
    # CREATE AN ARRAY OF *ALL* NATIVE HISTORY RECORDS 
    # REGARDLESS OF WHETHER OR NOT THEY ARE ANATHEMA
    # OR DUPLICATES. SAVE IT TO XML
    #
    $whereIwuz = get-location
    $GLOBAL:chronicle  = New-Object System.Collections.ArrayList
    write-host $nullstr
    write-Host -ForegroundColor Green "BEGINNING HARVEST..."
    $harvestCount = 0
    $subtotal = 0
    foreach ( $MatrixRow in $chronicalDirectoryMatrix ) {
        $whereIwuz = get-location
        $SourceDir = $MatrixRow[0]
        $messageTxt = $MatrixRow[1]
        write-host $nullstr
        write-Host -ForegroundColor Green $SourceDir
        write-Host -ForegroundColor Green $messageTxt
        write-host $nullstr
        set-Location  $SourceDir
        Get-ChildItem *.pshcmd.xml | %{ # {{{
            #
            # Gather commands and execution times into an ordinary array
            # because I want to leave as little as possible to chance.
            #
            try {
                $chronicleRecord = import-clixml $_
                $harvestCount = $GLOBAL:chronicle.add( $chronicleRecord )
                if ( 0 -eq $harvestCount % 512 )
                { # {{{
                    write-Host -ForegroundColor Gray "harvest count: $harvestCount"
                } # }}}
            }
            catch {
                $msg = "OUCH:<generateChronicle>[ $_.fullname ] "
                $msg += "APPEARS TO HAVE A PROBLEM"
                Write-Host
                Write-Host -ForegroundColor Red $msg
                Write-Host
            }
            $subtotal += 1
        } # }}}
        write-Host -ForegroundColor Gray "     subtotal: $subtotal"
        set-Location $whereIwuz
    }
    #
    # save chronicle to xml
    #
    #write-Host -ForegroundColor Green "SORTING AND SAVING CHRONICLE XML FILE"
    write-Host -ForegroundColor Green "SAVING CHRONICLE XML FILE"
    $xmltmpfile = $psDiaryLclD + '/' + $(usrHostTStampPid) + ".chronicle.xml"
    #$GLOBAL:chronicle | sort-object -property StartExecutionTime | Export-Clixml -Path $xmltmpfile
    Export-Clixml -Path $xmltmpfile -InputObject $GLOBAL:chronicle
    $m = $nullstr
    $m += $newline
    $m += $(dspc $xmltmpfile)
    $m += $newline
    write-host -foregroundcolor yellow $m
    Copy-Item $xmltmpfile $psChronicleFile
} # }}}


function generateDiary { # {{{
    $whereIwuz = get-location
    $uniq = @{}
    $raw = @()
    $duplicates = New-Object System.Collections.ArrayList
    $originalCount = $diary.count
    $GLOBAL:diary = @()
    $dontkeepthese = [regex] '(?six: \A xxr \s )'
    write-host $nullstr
    write-Host -ForegroundColor Green "BEGINNING HARVEST..."
    write-host $nullstr
    write-Host -ForegroundColor Green "ORIGINAL COUNT IS $originalCount"
    write-host $nullstr
    $harvestCount = 0
    write-host $nullstr
    write-Host -ForegroundColor Green "LOADING ARCHIVE ENTRIES..."
    write-host $nullstr
    set-Location  $psDiaryArchive
    Get-ChildItem *.pshcmd.xml | %{ # {{{
        #
        # Gather commands and execution times into an ordinary array
        # because I want to leave as little as possible to chance.
        #
        $xml2import = $_
        try {
            $histobj = import-clixml $xml2import
            $diaryRecord = [pscustomobject]@{
                         stamp = $histobj.StartExecutionTime.ticks.tostring()
                   commandline = $histobj.commandline
                        origin = $xml2import.name
            }
            if ( $diaryRecord.commandline -match $dontkeepthese )
            { # {{{
                $index = $duplicates.add($diaryRecord.origin)
            } # }}}
            else
            { # {{{
                $raw += ,( $diaryRecord )
            } # }}}

            $harvestCount += 1
            if ( 0 -eq $harvestCount % 512 )
            { # {{{
                write-Host -ForegroundColor Green "harvest count: $harvestCount"
            } # }}}
        }
        catch {
            $msg = "OUCH:<generateDiary>[ "
            $msg += $xml2import.fullname
            $msg += " ] "
            $msg += "APPEARS TO HAVE A PROBLEM"
            Write-Host
            Write-Host -ForegroundColor Red $msg
            Write-Host
        }
    } # }}}
    write-host $nullstr
    write-Host -ForegroundColor Green "LOADING RECENT ENTRIES..."
    write-host $nullstr
    set-Location  $psDiaryRecent
    Get-ChildItem *.pshcmd.xml | %{ # {{{
        #
        # Gather commands that have been executed since the last
        # time the diary was regenerated.
        #
        try {
            $histobj = import-clixml $_
            $diaryRecord = [pscustomobject]@{
                         stamp = $histobj.StartExecutionTime.ticks.tostring()
                   commandline = $histobj.commandline
                        origin = $_.name
            }

            if ( $diaryRecord.commandline -match $dontkeepthese )
            { # {{{
                $index = $duplicates.add($diaryRecord.origin)
            } # }}}
            else
            { # {{{
                $raw += ,( $diaryRecord )
            } # }}}

            $harvestCount += 1
            if ( 0 -eq $harvestCount % 512 )
            { # {{{
                write-Host -ForegroundColor Green "harvest count: $harvestCount"
            } # }}}
        }
        catch {
            $msg = "OUCH:<generateDiary>[ $_.fullname ] "
            $msg += "APPEARS TO HAVE A PROBLEM"
            Write-Host
            Write-Host -ForegroundColor Red $msg
            Write-Host
        }
    } # }}}
    if ( Test-Path -path $psDiaryArchive -PathType container ) {
        Move-Item *.pshcmd.xml $psDiaryArchive
    }
    set-Location $whereIwuz
    #
    # Sort by execution time
    #
    write-host $nullstr
    write-Host -ForegroundColor Green "BEGINNING SORT/CULL..."
    foreach( $diaryRecord in $raw | sort-object -property stamp )
    { # {{{
        #
        # weed out the duplicate commands
        # since they are sorted by date, the newest duplicate
        # will be in the diary at the relative
        # position where it was last invoked
        #
        try {
            $index = $uniq.add($diaryRecord.commandline, $diaryRecord)
        }
        catch {
            #
            # A duplicate key gets us here. So add the
            # origin file to the duplicates list to
            # move them out of the archive when we're done.
            # that way we're not constantly reprocessing
            # duplicates
            #
            $index = $duplicates.add($diaryRecord.origin)
            $uniq[ $diaryRecord.commandline ] = $diaryRecord
        }
    } # }}}
    $raw = @() # release the memory
    $GLOBAL:diary =  $uniq.values | sort-object -Property stamp
    $duplicates | %{
        $originName = $_ -replace ".*[/\\]", ""
        $originFullName = "$psDiaryRecent/$_"
        if ( Test-Path -path $originFullName -PathType leaf ) {
            try {
                move-item $originFullName $psDiaryTwins
            }
            catch {
                $m = @()
                $m+= $newline
                $m+= 'OUCH<generateDiary 67BB859D>: Move Duplicate Failed'
                $m+= $newline
                $m+= "move-item $originFullName $psDiaryTwins"
                $m+= $newline
                $m+= "**LOOKS LIKE IT FAILED. NOT SURE WHY**"
                $m+= $newline
                $errmsg = $m -join $newline
                write-Host -Foregroundcolor Red $errmsg
                move-item $originFullName "$psDiaryTwins/$(rndx32bit).$originFullName"
            }
        }
        else {
            $originFullName = "$psDiaryArchive/$_"
            if ( Test-Path -path $originFullName -PathType leaf ) {
                try {
                    move-item $originFullName $psDiaryTwins
                }
                catch {
                    $m = @()
                    $m+= $newline
                    $m+= 'OUCH<generateDiary 118ED662>: Move Duplicate Failed'
                    $m+= $newline
                    $m+= "move-item $originFullName $psDiaryTwins"
                    $m+= $newline
                    $m+= "**LOOKS LIKE IT FAILED. NOT SURE WHY**"
                    $m+= $newline
                    $errmsg = $m -join $newline
                    write-Host -Foregroundcolor Red $errmsg
                    move-item $originFullName "$psDiaryTwins/$(rndx32bit).$originFullName"
                }
            }
        }
    }
    $uniq = @{}  # release the memory
    #
    # dump array to flatfile
    #
    write-Host -ForegroundColor Green "...SORT COMPLETE"
    write-host $nullstr
    write-Host -ForegroundColor Green "DUMPING DIARY TEXT FILE"
    dumpDiary
    #
    # save diary to xml
    #
    write-Host -ForegroundColor Green "SAVING DIARY XML FILE"
    saveDiary
    if ( $originalCount -gt 0 ) {
        $dupcount = $originalCount - $diary.count
        if ( $dupcount -gt 0 ) {
            write-Host -ForegroundColor Green "$dupcount DUPLICATES ELIMINATED SINCE LAST GENERATION"
            write-host $nullstr
        }
    }
    $dupcount = $harvestCount - $diary.count
    write-Host -ForegroundColor Green "$dupcount DUPLICATES EXIST IN THE ARCHIVE"
    write-host $nullstr


} # }}}

function ssinit { # {{{
    #
    # initialize the Folderhash structure
    # from persistant storage.
    #
    $lclFolderHash = @{}
    $cronological = @{}
    $datePart = "(?sx: \d+ [.]   \d+    [.] \d+ )"
    $hostPart = "(?six: $localhost )"
    $pattern = [regex]"(?six: \A ( $datePart ) \s+ ( $hostPart ) \s+ ( \S+ ) \s* \z )"
    $pattern = [regex]"(?six: \A ( $datePart ) \s+ ( $hostPart ) \s+ ( .+ ) \z )"

    #
    # fetch entire cdJournal . filter for keys gathered on this host
    #
    if ( Test-Path -path $cdJournal -PathType leaf ) {
        Get-Content $cdJournal         |
        Out-String -stream             |
        Select-String $pattern         | foreach {
            #
            # EACH LINE OF cdJournal HAS THE FORM:
            #
            #           20120531.102712.7332716   PrimeLclHost  C:/nomad
            #
            $dstamp = $_ -replace $pattern, '$1'
            $dsplydir = $_ -replace $pattern, '$3'
            $dsplydir = dspc $dsplydir
            $lclFolderHash[ $dsplydir ] = $dstamp
        }

        $lclFolderHash.keys | foreach {
            #
            # switch the keys with the values
            #
            $cronological[ $lclFolderHash[ $_ ] ]=$_
        }

        $cronological.keys         |
        sort-object                | foreach {
            #
            # sort cronologically. get them all
            #
            $GLOBAL:FolderHash[ $cronological[ $_ ] ] = $_
        }
    }

} # }}}

function showDirectoryStack { # {{{
    Param (
        [Parameter(Position=0, Mandatory=$true)]
        [System.Object]$count,
        [Parameter(Position=1, Mandatory=$true)]
        [System.Object]$filter
    )
    #
    if ( $nullstr -eq $filter ) {
        $filter = '.'
    }
    #
    $cronological = @{}
    $forDisplay   = @()

    $FolderHash.keys | foreach {
        #
        # switch the keys with the values
        #
        $cronological[ $FolderHash[ $_ ] ] = $_
    }

    $cronological.keys |
    sort-object        | foreach {
        #
        # sort cronologically. take only the last count entries
        #
        $displayPath = $cronological[ $_ ]
        if ( $displayPath -match '(?six: \A // )' ) {
            $displayPath = '  ' + $displayPath
        }
        $forDisplay += $displayPath
    }

    $GLOBAL:RawUI = (get-host).UI.RawUI
    $RawUIOriginalColor = $RawUI.ForegroundColor

    $GLOBAL:RawUI.ForegroundColor = 'Gray'

    $forDisplay                        |
    Out-String -Stream                 |
    Select-String -pattern $filter     |
    select-Object -last $count         | foreach {
        write-Output $_
    }

    $GLOBAL:RawUI.ForegroundColor = $RawUIOriginalcolor

} # }}}

function ss { # {{{
    showDirectorystack 32 $args
}   # }}}

function sss { # {{{
    showDirectoryStack 128 $args
}   # }}}

function ssss { # {{{
    showDirectoryStack 256 $args
} # }}}

function ssall { # {{{
    showDirectoryStack 9999 $args
} # }}}

function pgy { # {{{
    #
    # start powergui for powershell debugging
    #
    pgui  @args
} # }}}

function setdir { # {{{
    $inpArg   = nSpc $(  bs2s $( qs $( [string]$args  ) ) )
    $inpArg   = $inpArg  -replace  $leadingWS, $nullstr           # trim leading white space
    $inpArg   = $inpArg  -replace  $trailingWS, $nullstr           # trim trailing white space

    $whereIwannaGo = $inpArg

    if (   $whereIwannaGo   -match  '(?six: \A  [.] \z ) ' ) {
        #
        # posh chokes on 'cd .' at comment: 'DlRECTORY DID NOT CHANGE'
        # but I could not figure out why because it only fails when
        # I am not running in powergui. This is a workaround.
        #
        $whereIwannaGo = currentDirectory
    }

    if ( $whereIwannaGo -ne $( currentDirectory ) -and $whereIwannaGo -ne $nullstr) {
        $iGottaGoBak = $false
        $whereIwuz = currentDirectory
        $errMsg = $nullstr
        if ( $whereIwuz -match $notAfileSystemRE ) {
            #
            #  if posh currentDirectory is Env: then
            #  Test-Path does not evaluate UNC paths properly.
            #  So...
            #  If current directory is Env: then go home first
            #
            $iGottaGoBak = $true # if the directory change fails then iGottaGoBak
            Set-Location $initialDirectory
        }
        if ( Test-Path -path $whereIwannaGo -PathType container ) {
            set-Location    $whereIwannaGo
            if ( $whereIwuz -ne $( currentDirectory ) ) {
                #
                #    CHANGE DIRECTORY SUCCEEDED
                #
                $GLOBAL:folderHash[ "$( currentDirDsply )" ] = mststamp
                #
                #    it looks like Add-Content does not check the file encoding
                #    prior to adding text. consequently the commented 'Add-Content'
                #    statements will bugger the directory history by adding plain
                #    ascii to a utf8 file. -- Go figure. I like add-content in
                #    principle so I am gonna leave these broken statements here
                #    till I figure out how to fix them by explicitely encoding
                #    my utf8 string into the destination.
                #
                #      $sout = "$( currentDirDsply )"
                #      Add-content -path $GLOBAL:cdTrail -value $sout
                #      $sout = "$( mststamp ) $( hostname ) $( currentDirDsply )"
                #      Add-Content -path $GLOBAL:cdJournal -value $sout
                #
                # The redirection works, but, I rather use Add-Content cause
                # it is more in keeping with powershell style.
                #
                $sout = "$( currentDirDsply )"
                write-Output $sout >> $cdTrail
                $sout = "$( mststamp ) $localhost $( currentDirDsply )"
                write-Output $sout >> $cdJournal
            }
            else {
                #
                #    DIRECTORY DID  NOT CHANGE
                #
                $o1 = $fso.getfolder( $whereIwuz )
                $n1 = $o1.path
                $o2 = $fso.getfolder( $whereIwannaGo )
                $n2 = $o2.path
                if ( $n1 -ne $n2 ) {
                    $errMsg  += "OUCH<setdir>:[ $whereIwannaGo ] "
                    $errMsg  += "APPEARS TO BE A DIRECTORY, BUT "
                    $errMsg  += "DIRECTORY DID NOT CHANGE"
                }
            }
        }
        else {
            $errMsg += "OUCH<setdir>:[ $whereIwannaGo ] "
            $errMsg += "DOES NOT APPEAR TO BE A DIRECTORY"
        }

        if ( $errMsg ) {
            $errMsg = $newline * 3 + $errMsg + $newline * 3
            write-Host -ForegroundColor Red $errMsg
            if   ( $iGottaGoBak ) {
                Set-Location $whereIwuz
                $iGottaGoBak = $false
            }
        }
    }

} # }}}

function mmm { #{{{
     #
     #  go home
     #
     setdir $GLOBAL:mmm
} # }}}

function nukearchiology { #{{{
    #
    # I keep timestamped copies of...
    #   history
    #   this file
    #   transcripts
    #   password files
    #
    # This creates a lot of junk in my directory hierarchy
    # but is very good for tracking things when something
    # goes horribly wrong. But, once something gets old enough
    # it is not so useful. So, this it what I use to clean
    # up all the bread crumbs I have left in the forest to
    # find my way back home.
    #
    $nukeCandidatex = $expiredFiles
    if (   Test-Path -path $nukeCandidateX -PathType leaf ) {
        Get-Content  $nukecandidatex |
        sort-object  -unique | foreach  {
            $artifact = $_
            if ( Test-Path -path $artifact -PathType leaf ) {
                $fObj = get-Item $artifact
                $tMod = "{0:yyyyMMdd.HHmmss}" -f $fObj.LastwriteTime
                write-Host -ForegroundColor Magenta "NUKING -> [ $tMod $artifact ]"
                Remove-Item $artifact
            }
        }
        Remove-Item $nukeCandidatex
    }
} #}}}

function rdtdesales { # {{{
    $originalLocation = $PWD
    setdir "$bin"
    invoke-expression -command  "$sys64L/mstsc.exe /f deSalesWifiNomad.rdp"
    Set-Location $originalLocation
} # }}}

function utf8 { # {{{
    #
    # UTF8 ENCODE THE DESIRED DATA
    #
    $argstring = nSpc $( bs2s $( qs $( [string]$args ) ) )
    if ( $nullstr -eq $argstring ) {
        #
        # null argument implies input from stdin
        #
        $w2s = "$sundryTmpD/"
        $w2s += "$( usrHostTStampPid ).$(  rnd32bit ).UTF8.txt"
        $where2save = bs2s $w2s
        $input | Out-String | set-Content -encoding UTF8   $where2save
        write-Output "UTF8 CONTENT IS IN: [ $where2save ]"
    }
    else{
        $elementAvailable = $false
        $elementObjects = Get-item -force $( nspc $argstring ) -ErrorAction silentlyContinue
        foreach ( $e2check in $elementobjects ) {
            if ( $null -ne $e2check ) {
                $elementAvailable = $e2check.mode -notmatch "d"
            }
        }
        if ( $elementAvailable ) {
            Get-Item -force $( nSpc $argstring  ) | foreach {
                # {{{
                    # handle any wildcard expansion that may
                    # be embedded in '$argstring'
                # }}}
                $entry = $_
                if ( $entry.mode -notmatch "d" ) {
                    $fullPathName = $( bs2s ($entry.FullName).ToString() )
                    $w2s = "$sundryTmpD/"
                    $w2s += "$( usrHostTStampPid ).$( rnd32bit ).MAYBE-NOT-UTF-8.txt"
                    $where2save = bs2s $w2s
                    Move-item $fullPathName $where2save
                    get-content $where2save | Set-Content -encoding UTF8 $fullPathName
                    $infomsg = "ORIGINAL [ $fullPathName ] CONTENT PRESERVED IN: [ $where2save ]"
                    write-Host -Foregroundcolor Green $infomsg
                }
                else {
                    $errormessage = "OUCH<utf8>:[ $entry ] IS A DIRECTORY."
                    write-Host -ForegroundCo1or Red $errormessage
                }
            }
        }
        else   {
            $errormessage     =  "OUCH<utf8>:[ $argstring ] "
            $errormessage    +=  "DOES NOT APPEAR TO MATCH ANY EXISTING file."
            Write-Host -Foregroundcolor Red $errormessage
        }
    }
} # }}}

function sevenbit { # {{{
    #
    # ASCII ENCODE THE DESIRED DATA
    #
    $argstring = nSpc $( bs2s $( qs $( [string]$args ) ) )
    if ( $nullstr -eq $argstring ) {
        #
        # null argument implies input from stdin
        #
        $w2s = "$sundryTmpD/"
        $w2s += "$( usrHostTStampPid ).$(  rnd32bit ).ascii.txt"
        $where2save = bs2s $w2s
        $input | Out-String | set-Content -encoding ascii   $where2save
        dos2unix $where2save
        write-Output "7-BIT ASCII CONTENT IS IN: [ $where2save ]"
    }
    else{
        $elementAvailable = $false
        $elementObjects = Get-item -force $( nspc $argstring ) -ErrorAction silentlyContinue
        foreach ( $e2check in $elementobjects ) {
            if ( $null -ne $e2check ) {
                $elementAvailable = $e2check.mode -notmatch "d"
            }
        }
        if ( $elementAvailable ) {
            Get-Item -force $( nSpc $argstring  ) | foreach {
                # {{{
                    # handle any wildcard expansion that may
                    # be embedded in '$argstring'
                # }}}
                $entry = $_
                if ( $entry.mode -notmatch "d" ) {
                    $fullPathName = $( bs2s ($entry.FullName).ToString() )
                    $w2s = "$sundryTmpD/"
                    $w2s += "$( usrHostTStampPid ).$( rnd32bit ).MAYBE-UTF-8.txt"
                    $where2save = bs2s $w2s
                    Move-item $fullPathName $where2save
                    get-content $where2save | Set-Content -encoding ascii $fullPathName
                    dos2unix $fullPathName
                    $infomsg = "ORIGINAL [ $fullPathName ] CONTENT PRESERVED IN: [ $where2save ]"
                    write-Host -Foregroundcolor Green $infomsg
                }
                else {
                    $errormessage = "OUCH<sevenbit>:[ $entry ] IS A DIRECTORY."
                    write-Host -ForegroundCo1or Red $errormessage
                }
            }
        }
        else   {
            $errormessage     =  "OUCH<sevenbit>:[ $argstring ] "
            $errormessage    +=  "DOES NOT APPEAR TO MATCH ANY EXISTING file."
            Write-Host -Foregroundcolor Red $errormessage
        }
    }
} # }}}

function parsevimargs { # {{{
    #
    # {{{
    #
    # This long winded routine is intended to overcome some problems associated with
    # the way the powershell mangles command-line arguments and the way microsoft
    # is enamored of spaces in file names. It should not be this hard.
    #
    # Powershell will try to figure out what you "mean"
    #   so...
    #
    #   given 3 files:
    #                 1   tv.pl
    #                 2   op_imp_m3.ini
    #                 3   very silly name
    #
    #   parsevimargs tv.pl,op_imp_m3.ini,very silly name
    #
    #   ...will provide 3 arguments to parsevimargs. they are:
    #
    #                 1   tv.pl op_imp_m3.ini very
    #                 2   silly
    #                 3   name
    #
    #   ...not what most would expect.
    #
    #   since I expect them to be file names, I re-examine the components and test
    #   to see if the files exist and make a marginally good faith effort
    #   to create a list of existing file names in a
    #   meta-load file (a file that contains a list of files to load.)
    #   vimrc uses the single meta-load file to load the files.
    #   into its window buffers
    #
    # }}}
    #
    $prefix = UidTStampRnd32
    #
    #   $vimMetaLoad is a temp file that will contain a
    #   linefeed separated list of filenames that vim
    #   should try to load into itz buffers
    #
    $vimMetaLoad = bs2s "$sundryTmpD/$prefix.vimtmp.daolatem"
    $currentdirectory = bs2s $( pwd )
    #
    # first line of the metaload file is the current directory
    #
    Set-Content -path $vimMetaLoad -Encoding UTF8 -value $currentdirectory

    if ( 1 -eq $args.count ) { # {{{
        $arg00 = [string] $args[ 0 ]
        if ( "-" -eq $arg00 ) {
            #
            # this semantic means:
            #           1          look for piped input
            #           2          redirect that input to a temp file
            #           3          open that temp file
            #
            $vimtmp =    bs2s "$sundryTmpD/$prefix.vimtmp.txt"
            $input | out-file $vimtmp -encoding utf8
            $args[ 0 ] = $vimtmp
            if ( $vvvIsVerbose ) {
                write-Host -ForegroundColor Green $( dspc $vimtmp )
            }
        }
    } # }}}


    $count = 0
    $expandedArgs = @{}
    foreach(    $subcomponent in $args ) { # {{{ split argument list on whitespace
        #
        #   work my way through all the supplied
        #   command line arguments and put them
        #   in a structure I can control ( a hash )
        #   I will create text hash keys that will
        #   easily sort into the correct order.
        #   buildkey creates a decimal number based
        #   key that is zero-padded to 4 places.
        #   so, '$expandedargs' indexes are...
        #
        #    0000
        #    0001
        #    0002
        #
        #
        #
        #    0DDD
        #
        #   in the event that more than 9999 arguments
        #   are supplied on the command-1ine, an un-handled
        #   error will occur. I don`t think powershell
        #   allows that many arguments, so it will barf
        #   before I do.
        #
        $stringcomponent = [string] $subcomponent
        if ( $stringcomponent -match "\s+" ) {
            #
            # I found whitespace within an element
            # of the args array
            #
            $stringcomponent -split "\s+" | foreach {
                $expandedArgs[ $( buildKey $count ) ] = $_
                $count = $count + 1
            }
        }
        else    {
            $expandedArgs[ $( buildKey $count ) ] = $stringcomponent
            $count = $count + 1
        }
    } # }}}

    #
    # reason for @(  $expandedArgs.keys | sort-object ) {{{
    #
    # unless I wrap '@( )' around '$expandedArgs.keys | sort-object'
    # the use-case of a single argument without embedded spaces will fail.
    # '$orderedKeys.count' will be undefined   (in this use-case) because
    # 'sort-object' returns an array *ONLY* if more than 1 entry
    # gets sorted [  go figure ]. otherwise it returns a string. (and string has
    # no 'count' property)
    #
    # }}}
    #
    $orderedKeys = @( $expandedArgs.keys | sort-object )
    $failsafe = 0
    if ( $orderedKeys.count ) {
        #
        # 1 or more arguments were provided on command line
        #
        $startPos = 0
        :out while ( $failsafe++ -lt 10000 ) { # {{{ Failsafe exit
            # {{{
            #  This loop is moderately tricky.
            #  It is pretty well tested
            #  I do not think there is any way to
            #  get stuck inside it. But the failsafe
            #  makes sure.
            #
            #  EXECUTIVE SUMMARY: work my way through the argument list until
            #  I find a plausible way to split them up such that all the arguments
            #  actually refer to readable files ( as opposed to partial file names )
            # }}}
            $errMsg = $nullstr
            $elementname = $nullstr
            for ( $ix = $startPos; $ix -lt $orderedKeys.count; $ix++ ) {
                # {{{ might come back here with new startpos
                $index = $orderedKeys[ $ix ]
                #
                # get the next file-name component
                #
                $suffix = $expandedArgs[ $index ]
                #
                # append it to the previous
                #
                $elementname = "$elementname $suffix"
                #
                # trim the leading whitespace
                #
                $elementname = $elementname -replace $leadingWS, $nullstr
                $elementAvailable = $false
                $elementObjects = $null
                # {{{
                # continue if element does not exist. otherwise, get the object.
                #
                # The 'elementAvailable' was murderous to get to
                # because hidden files are even hidden from "test-path".
                #
                # The only way to see hidden entries (files/directories) is either
                # "get-childitem -force" or "get-item -force". But, both of
                # those generate errors for elements which are not found.
                #
                # 'try-catch' does not supress these errors.
                #
                # 'trap' does not supress these errors.
                #
                # '-ErrorAction silentlyContinue' did the trick.
                #
                # The acid test was having wild carded hidden files and $profile
                # as arguments to the same "vvv" invocation. It should not be this hard.
                # }}}
                $elementobjects = Get-Item -force $( nspc $elementname ) -ErrorAction silentlycontinue
                foreach ( $e2check in $elementobjects ) {
                    # {{{
                    #  $elementObjects will be $null on a partial path
                    #  if that happens then element remains unavalible
                    #  skip the directories that match wild card specs
                    # }}}
                    if ( $null -ne $e2check ) {
                        $elementAvailable = $e2check.mode -notmatch "d"
                    }
                }
                if ( $elementAvailable ) {
                    # {{{ Process an elementpath because powershell says it exists
                    #
                    # a file exists with the name '$elementname'
                    #
                    Get-Item -force $( nspc  $elementname ) | foreach {
                        # {{{
                        #   handle any wildcard expansion that may
                        #   be embedded in '$elementname1
                        # }}}
                        $entry = $_
                        if ( $entry.mode -notmatch "d" ) {
                            $fullPathName = $( bs2s ($entry.FullName).ToString()                                         )
                            # {{{
                            # append the next filename to meta-load file
                            # }}}
                            Add-Content -path $vimMetaLoad -Encoding utf8 -value $( dSpc $fullPathName )
                        }
                    }
                    # {{{
                    #  I am done with this element, so make sure
                    #  it does not become the prefix for the next element
                    # }}}
                    $elementname = $nullstr
                    $startPos = $ix + 1
                    # {{{
                    # any previously build up error message  from the 'else'
                    # clause (below) needs to be cleared
                    # }}}
                    $errMsg = $nullstr
                    if ( $ix -eq $orderedKeys.count - 1 ) {
                        # {{{
                        # everything in the argument list has been processed
                        # }}}
                        break out
                    }
                    # }}}
                }
                else { # {{{ set up a retry from a different startpos or exit altogether
                    # {{{
                    # '$elementname1 was not found, but it might be the prefix
                    # for another $elementname that will be  found, so save
                    # the error message. But, do not display it yet
                    # }}}
                    $errMsg = $errMsg + $newline + "NOT FOUND:: [ $elementname ]"
                    if ( $ix -eq $orderedKeys.count - 1 ) {
                        # {{{
                        #  I have reached the end of all imput arguments
                        #  so, this argument really has no match on the filesystem
                        # }}}
                        write-Host -ForegroundColor Red $errMsg
                        # {{{
                        #  move the cursor to the next commandline argument
                        # }}}
                        $startPos = $startPos + 1
                        # {{{
                        #  I already printed the error message, so, clear it.
                        # There is no need to repeat myself.
                        # }}}
                        $errMsg = $nullstr
                    }
                    if ( $startPos -eq $orderedKeys.count ) {
                        # {{{
                        # the cursor has reached the end of all imput arguments
                        # so, now we are really finished. Asta.
                        # }}}
                        write-Host -ForegroundColor Red $errMsg
                        break out
                    }
                } # }}}
                # }}}
            }
        }  # }}}
    }

    if ( $vvvIsVerbose ) {
        write-Host -ForegroundColor Green $( dspc $vimMetaLoad )
    }
    $vimMetaLoad
} # }}}

function tfil { # {{{
    # return a unique name for a temp file
    $prefix = UidTStampRnd32
    "$sundryTmpD/$prefix.vimtmp.txt"
} # }}}

function yyy { # {{{
    $slickedit = 'C:\PROGRA~1\SLICKE~1.2\win\vs.exe'
    write-Host -ForegroundColor Green
    $errmsg = "** INVOKING SLICKEDIT ON $args **"
    write-Host -ForegroundColor Green $errmsg
    write-output $nullstr
    if ( $nullstr -ne $slickedit ) {
        foreach( $component in $args ) {
            $expression2invoke = $slickedit + $space + $(sqlquote $component)
            invoke-Expression $expression2invoke
        }
    }
} # }}}

function uutty { # {{{
    $puttyimage = 'C:\PROGRA~1\Putty\putty.exe'
    write-Host -ForegroundColor Green
    $errmsg = "** INVOKING PUTTY -LOAD ON $args **"
    write-Host -ForegroundColor Green $errmsg
    write-output $nullstr
    if ( $nullstr -ne $puttyimage ) {
        foreach( $component in $args ) {
            $expression2invoke = $puttyimage + $space + '-load' + $space + $(sqlquote $component)
            invoke-Expression $expression2invoke
        }
    }
} # }}}

function ttytiv { # {{{
    uutty pssva062
} # }}}

function ttyomn { # {{{
    uutty pvlva614
} # }}}

function tgw { # {{{
    657,658,659,660,661| %{putty pvsva$_}
} # }}}

function tivomnx2 { # {{{
    uutty pvlva614
    uutty pvlva614
    uutty pssva062
    uutty pssva062
} # }}}

function gw13 { # {{{
    uutty pvsva013
} # }}}

function gw14 { # {{{
    uutty pvsva014
} # }}}

function gw15 { # {{{
    uutty pvsva015
} # }}}

function vimfallback { # {{{
    write-output $nullstr
    $errmsg = "OUCH<vimfallback>:[ ** CURRENT HOSTNAME APPEARS TO BE $localhost ** ]"
    write-Host -ForegroundColor Red $errmsg
    write-output $nullstr
    write-Host -ForegroundColor Red
    $errmsg = "** [ $gooeyvimexe ] DOES NOT APPEAR TO BE A FILE ON $localhost **"
    write-Host -ForegroundColor Red $errmsg
    write-output $nullstr
    $slickedit = 'C:\PROGRA~1\SLICKE~1.2\win\vs.exe'
    if ( $nullstr -ne $slickedit ) {
        foreach( $component in $args ) {
            $expression2invoke = $slickedit + $space + $(sqlquote $component)
            invoke-Expression $expression2invoke
        }
    }
} # }}}

function starteclipse { # {{{
    if ( Test-Path -path $eclipseimage -PathType leaf ) {
        $originalLocation = $PWD
        setdir $eclipseDir
        #
        # eclipse will not run with any of...
        #
        #    $pfxL/java/jdkl.6.0_18/jre/bin
        #    $pfxL/java/jdkl.6.0_18/bin
        #    $pfdL/java/jdkl.6.0_18/jre/bin
        #    $pfdL/java/jdkl.6.0_18/bin
        #
        # ...in the path, unless the vm is explicitely called out.
        #    go figure.
        #
        invoke-Expression "$eclipseimage -vm $eclipsevm"
        Set-Location $originalLocation
    }
    else {
        $errMsg = $newline + "NOT FOUND:: [ $eclipseimage ]"
        write-Host -ForegroundColor Red $errmsg
        write-output $nullstr
    }
} # }}}

function vvv { # {{{
    # {{{
    # call gvim in a round about way.
    # the striking thing about this line
    # is that in order to make sure
    # parsevimargs gets the same arguments
    # as vvv you must:
    # parsevimargs @args
    # because:
    # parsevimargs $args
    #
    # *WILL NOT WORK*
    #
    # I suspect this is a deep dark secret
    # of powershell. I did not see it documented
    # anywhere, but I thought I`d give it a try
    # and @args *does* work. it is best not to
    # argue with success.
    # }}}

    $iGottaGoBak = $false
    if ( ( currentDirectory      ) -match $notAfileSystemRE    ) {
        $iGottaGoBak = $true
        Push-Location $initialDirectory
    }


    if ( Test-Path -path $gooeyvimexe -PathType leaf )         {
        $metaFile = $( $input | parsevimargs @args )
        $cmd2run = "$gooeyvimexe -u $vimprofile $metaFile"
        write-Host -ForegroundColor DarkYellow $cmd2run
        Invoke-Expression $cmd2run
    }
    else {
        write-output $nullstr
        $errmsg = "OUCH<gvimfallback>:[ ** CURRENT HOSTNAME APPEARS TO BE $localhost ** ]"
        write-Host -ForegroundColor Red $errmsg
        write-output $nullstr
        write-Host -ForegroundColor Red
        $errmsg = "** [ $gooeyvimexe ] DOES NOT APPEAR TO BE A FILE ON $localhost **"
        write-Host -ForegroundColor Red $errmsg
        write-output $nullstr
        vv @args
    }

    if ( $iGottaGoBak ) {
        Pop-Location
        $iGottaGoBak = $false
    }
} # }}}

function vvvr { # {{{
    $iGottaGoBak = $false
    if ( ( currentDirectory ) -match $notAfileSystemRE ) {
        $iGottaGoBak = $true
        Push-Location $initialDirectory
    }


    if ( Test-Path -path $gooeyvimexe -PathType leaf ) {
        $metaFile = $( $input | parsevimargs @args )
        $cmd2run = "$gooeyvimexe -u $vimprofile -R $metaFile"
        write-Host -ForegroundColor DarkYellow $cmd2run
        Invoke-Expression $cmd2run
    }
    else {
        write-output $nullstr
        $errmsg = "OUCH<gvimfallback>:[ ** CURRENT HOSTNAME APPEARS TO BE $localhost ** ]"
        write-Host -ForegroundColor Red $errmsg
        write-output $nullstr
        write-Host -ForegroundColor Red
        $errmsg = "** [ $gooeyvimexe ] DOES NOT APPEAR TO BE A FILE ON $localhost **"
        write-Host -ForegroundColor Red $errmsg
        write-output $nullstr
        vvr @args
    }

    if ( $iGottaGoBak ) {
        Pop-Location
        $iGottaGoBak = $false
    }
} # }}}

function vv { # {{{
    $iGottaGoBak = $false
    if ( ( currentDirectory ) -match $notAfileSystemRE ) {
        $iGottaGoBak = $true
        Push-Location $initialDirectory
    }

    if ( Test-Path -path $crunchyvimexe -PathType leaf ) {
        $metaFile = $( $input | parsevimargs @args )
        $d = $doublequote
        $cmd2run =  "start-process -windowstyle normal -filepath " +
                    "$sys32/windowsPowerShell/v1.0/powershell.exe -argumentlist " +
                    "${d}-noprofile -command $crunchyvimexe ${d}${d}-u $vimprofile $metaFile${d}${d}${d}"

        write-Host -ForegroundColor DarkYellow $cmd2run
        get-content $metafile
        Invoke-Expression -command $cmd2run
    }
    else {
        vimfallback @args
    }

    if ( $iGottaGoBak ) {
        Pop-Location
        $iGottaGoBak = $false
    }
} # }}}

function vvr { # {{{
    $iGottaGoBak = $false
    if ( ( currentDirectory ) -match $notAfileSystemRE ) {
        $iGottaGoBak = $true
        Push-Location $initialDirectory
    }


    if ( Test-Path -path $crunchyvimexe -PathType leaf ) {
        $metaFile = $( $input | parsevimargs @args )
        $d = $doublequote
        $cmd2run =  "start-process -windowstyle maximized -filepath " +
                    "$sys32/windowsPowerShell/v1.0/powershell.exe -argumentlist " +
                    "${d}-noprofile -command $crunchyvimexe ${d}${d}-u $vimprofile -R $metaFile${d}${d}${d}"

        write-Host -ForegroundColor DarkYellow $cmd2run
        get-content $metafile
        Invoke-Expression -command $cmd2run
    }
    else {
        vimfallback @args
    }

    if ( $iGottaGoBak ) {
        Pop-Location
        $iGottaGoBak = $false
    }
} # }}}

function vvd { # {{{
    $iGottaGoBak = $false
    if ( ( currentDirectory ) -match $notAfileSystemRE ) {
        $iGottaGoBak = $true
        Push-Location $initialDirectory
    }


    if ( Test-Path -path $crunchyvimexe -PathType leaf ) {
        $metaFile = $( $input | parsevimargs @args )
        $diffMeta = $metaFile + ".ffidmiv"
        move-item $metaFile $diffMeta
        $metaFile = $diffMeta
        $d = $doublequote
        $cmd2run =  "start-process -windowstyle maximized -filepath " +
                    "$sys32/windowsPowerShell/v1.0/powershell.exe -argumentlist " +
                    "${d}-noprofile -command $crunchyvimexe ${d}${d}-u $vimprofile -R $metaFile${d}${d}${d}"

        write-Host -ForegroundColor DarkYellow $cmd2run
        get-content $metafile
        Invoke-Expression -command $cmd2run
    }
    else {
        vimfallback @args
    }

    if ( $iGottaGoBak ) {
        Pop-Location
        $iGottaGoBak = $false
    }
} # }}}

function absrmpw { # {{{
    $image = s2bs( "$bin/AnyBizSoft/PDFPasswordRemover/PDFPasswordRemover.exe" )
    invoke-expression -command $image
} # }}}

function  __displayDiary { # {{{
    $fnCtl = $args;
    if ( 1 -eq $fnCtl.count )
    { # {{{
        #
        # the caller ( one of hh, hhh, hhhh )
        # did not receive any filter on the
        # command line. so create the filter
        # that accepts everything.
        #
        $fnCtl = ( ".", $args[ 0 ] )
    } # }}}
    $limit = $diary.count
    if ( $limit -gt 0 )
    { # {{{
        $limit -= 1
        $firstCmd = $limit - $fnCtl[ 1 ]
        if ( 0 -gt $firstCmd )
        { # {{{
            $firstCmd = 0
        } # }}}
        foreach( $i in $firstCmd .. $limit )
        { # {{{
            $cmdID = $i
            $cmdLine = $diary[ $i ].commandline
            #
            # split multi-line commands into an array
            # so that I can better control the display
            # of those lines.
            #
            $cmdLineArray = $cmdLine -split '[\n\r]+'
            if ( $cmdLine -match $fnCtl[ 0 ] )
            { # {{{
                foreach ( $entry in $cmdLineArray )
                { # {{{
                    if ( $fnCtl[ 0 ] -notmatch '\Ahh' -and  $entry -match '\Ahh' )
                    { # {{{
                        continue      # dont include other hh commands
                    } # }}}
                    $displayline = "H{0:D6}: {1}" -f $i, $entry
                    write-Output $displayline
                } # }}}
            } # }}}
        } # }}}
    } # }}}
    else
    { # {{{
        $msg = "OUCH:< __displayDiary >[ $(qs $args) ] "
        $msg += "YOUR DIARY IS EMPTY. GET BUSY YOU JUGHEAD."
        Write-Host -ForegroundColor Red $msg
    } # }}}
} # }}}

function ForbidDiaryCommand { # {{{
    try {
        if ( 1 -gt $args.count ) {
            Write-Host $newline
            $msg = "OUCH:< ForbidDiaryCommand >[ " + $args[ 0 ] + " ] "
            $msg += "ONE OR MORE ARGUMENTS ARE REQUIRED"
            Write-Host -ForegroundColor Red $msg
            Write-Host $newline
            throw
        }
        $argLimit = $args.count - 1
        $validationPattern = [regex]'(?six:  \A H [0]* (\d+) : \Z)';
        foreach( $argIndex in 0 .. $argLimit) {
            if ( -not ( $args[ $argIndex  ] -match $validationPattern ) ) {
                Write-Host $newline
                $msg = "OUCH:< ForbidDiaryCommand >[ " + $args[ 0 ] + " ] "
                $msg += "COMMAND ARGUMENT INVALID. MUST MATCH: (?six:  \A H [0]* (\d+) : \Z)"
                Write-Host -ForegroundColor Red $msg
                Write-Host $newline
                throw
            }
            $diaryArrayIndex = $args[ $argIndex ] -replace $validationPattern, '$1'
            $diaryArrayLimit = $diary.count - 1
            if ( $diaryArrayLimit -lt 0 ) {
                $diaryArrayLimit = 0
            }
            if ( -not ( $diaryArrayIndex -In 0 .. $diaryArrayLimit ) ){
                Write-Host $newline
                $msg = "OUCH:< ForbidDiaryCommand >[ " + $diaryArrayIndex + " ] "
                $msg += "COMMAND ARGUMENT OUT OF RANGE (0 - $diaryArrayLimit)."
                Write-Host -ForegroundColor Red $msg
                Write-Host $newline
                throw
            }
        }
        foreach( $argIndex in 0 .. $argLimit) {
            $diaryArrayIndex = $args[ $argIndex ] -replace $validationPattern, '$1'
            $diaryArrayLimit = $diary.count - 1
            if ( $diaryArrayLimit -lt 0 ) {
                $diaryArrayLimit = 0
            }
            $commandline =  $diary[ $diaryArrayIndex ].commandline
            $origin =  $diary[ $diaryArrayIndex ].origin
            Write-Host $newline
            Write-Host -ForegroundColor yellow $origin
            Write-Host $newline
            Write-Host -ForegroundColor yellow $commandline
            Write-Host $newline
            $originName = $origin -replace ".*[/\\]", ""
            $originFullName = "$psDiaryRecent/$originName"
            if ( Test-Path -path $originFullName -PathType leaf ) {
                move-item $originFullName $psDiaryAnathema
            }
            $originFullName = "$psDiaryArchive/$originName"
            if ( Test-Path -path $originFullName -PathType leaf ) {
                move-item $originFullName $psDiaryAnathema
            }
        }
    }
    catch {
        Write-Host $newline
        $msg = "OUCH:< Function ForbidDiaryCommand > args == [ " + $args + " ] "
        $msg += "EXCEPTION THROWN"
        Write-Host -ForegroundColor Red $msg
        Write-Host $newline
    }
} # }}}

function InvokeDiaryCommand { # {{{
    $pattern = "(?six:  \A H [0]* (\d+) : )";
    $id = $args[ 0 ] -replace $pattern, '$1'
    $commandline =  $diary[ $id ].commandline
    $scriptblk = $ExecutionContext.InvokeCommand.NewScriptBlock( $commandline )
    Write-Host $newline
    Write-Host -ForegroundColor yellow $commandline
    Write-Host $newline
    & $scriptblk
} # }}}

function hh { # {{{
    #
    # __displayDiary will accept 1 or 2 arguments.
    # if it receives 2 arguments then 1 of them is passed
    # through from hh, hhh, or hhhh. That argument is a
    # string which __displayDiary filters on.
    #
    # hh and friends will append the number parameter either
    # on to an empty array or onto an array that contains
    # a filter string.
    #
    $args += 64;
    __displayDiary @args
} # }}}

function hhh { # {{{
    $args += 256;
    __displayDiary @args
} # }}}

function hhhh { # {{{
    $args += $diary.count
    __displayDiary @args
} # }}}

function drecall { # {{{
    Param (
    [Parameter(Position=0, Mandatory=$false)]
    [int]$count=16,
    [Parameter(Position=1, Mandatory=$false)]
    [string]$filter="."
    )
    __displayDiary @($filter, $count)
} # }}}

function dwin { # {{{
    vvvr $psDiaryTxtFile
} # }}}

function reboot { # {{{
    Powershell_Exit_Processing
    restart-computer -force
} # }}}

function clenv1 { # {{{
    #
    # environment for clion
    #
    pathrm C:/unx/msys64/mingw64/bin
    pathrm C:/unx/msys64/usr/bin
    pathrm $git
    if (Test-Path variable:global:PYTHONHOME) {
        Remove-Item variable:global:PYTHONHOME
    }
    if (Test-Path variable:global:PYTHONHOMEL) {
        Remove-Item variable:global:PYTHONHOMEL
    }
    if (Test-Path Env:/PYTHONHOME) {
        Remove-Item Env:/PYTHONHOME
    }
    if (Test-Path Env:/PYTHONHOMEL) {
        Remove-Item Env:/PYTHONHOMEL
    }
    $GLOBAL:consoleTitleMode = 'ClEnvPath'
    $RawUI = (get-host).UI.RawUI
    $RawUI.backgroundcolor = 'Green'
    SetConsolewindowTitle
    clear-Host

} #}}}

function clpath { # {{{
    $heredoc = @"
    PATH
    $din
    $bin
    $sqlite
    $adb
    $sysinternals
    $bin/cmake/bin
    $bin/JetBrains/PyCharm/bin
    $bin/JetBrains/CLion/bin
    $bin/JetBrains/CLion1.2.4/bin
    $bin/7Zip
    $bin/Vim/vim74
    $bin/PuTTY
    $bin/jdk1.8.0.40/bin
    $abwL/eclipse
    $abwL/sdk/platform-tools
    $pfdL/firefox
    $spc/bin
    $sby/bin
    $sby/site/bin
    $pyd
    $pyd/Lib
    $pyd/Tools/Scripts
    $pyd/Scripts
    $win/system32
    $win
    $win/System32/Wbem
    $win/system32/WindowsPowerShell/v1.0
    $pfdL/QuickTime/QTSystem
    $pfxL/Common Files/Microsoft Shared/Windows Live
    $pfdL/Common Files/Microsoft Shared/Windows Live
    $win/system32
    $win
    $win/System32/Wbem
    $win/System32/WindowsPowerShell/v1.0
    $pfdL/Windows Live/Shared
    $bin/WIDCOMMbluetooth
    $bin/WIDCOMMbluetooth/syswow64
    $pfdL/Microsoft SQL Server/100/Tools/Binn
    $pfxL/Microsoft SQL Server/100/Tools/Binn
    $pfxL/Microsoft SQL Server/100/DTS/Binn
    $pfdL/Windows Live/Writer
    $pfdL/Calibre2
    $xbn
    $bin/qchat
"@
    ingest $heredoc
}   # }}}

function pypath { # {{{
    $heredoc = @"
    PYTHONPATH
    $bin
    $bin/PyCharmPro/debug-eggs/pycharm-debug-py3k
    $bin/JetBrains/PyCharm/debug-eggs/pycharm-debug-py3k
"@
    ingest $heredoc
} # }}}

function clenv { # {{{
    #
    # environment for clion
    #
    if (Test-Path variable:global:PYTHONHOME) {
        Remove-Item variable:global:PYTHONHOME
    }
    if (Test-Path variable:global:PYTHONHOMEL) {
        Remove-Item variable:global:PYTHONHOMEL
    }
    if (Test-Path Env:/PYTHONHOME) {
        Remove-Item Env:/PYTHONHOME
    }
    if (Test-Path Env:/PYTHONHOMEL) {
        Remove-Item Env:/PYTHONHOMEL
    }
    clpath
    $GLOBAL:consoleTitleMode = 'ClEnvPath'
    $RawUI = (get-host).UI.RawUI
    $RawUI.backgroundcolor = 'Green'
    SetConsolewindowTitle
    clear-Host
    Invoke-Expression $bin/jetbrains/CLion1.2.4/bin/clion64.exe
} #}}}

function vsenv { # {{{
    $envChain = @()

    $envChain += @"
    DevEnvDir
    $bin/VisualStudio14_0/Common7/IDE
"@

    $envChain += @"
    FrameworkDIR32
    $win/Microsoft.NET/Framework
"@

    $envChain += @"
    FrameworkDIR64
    $win/Microsoft.NET/Framework64
"@

    $envChain += @"
    INCLUDE
    $bin/VisualStudio14_0/VC/INCLUDE
    $pfdL/Windows Kits/10/include/10.0.10240.0/ucrt
    $pfdL/Windows Kits/NETFXSDK/4.6/include/um
    $pfdL/Windows Kits/10/include/10.0.10240.0/shared
    $pfdL/Windows Kits/10/include/10.0.10240.0/um
    $pfdL/Windows Kits/10/include/10.0.10240.0/winrt
"@

    $envChain += @"
    LIB
    $bin/VisualStudio14_0/VC/LIB/amd64
    $pfdL/Windows Kits/10/lib/10.0.10240.0/ucrt/x64
    $pfdL/Windows Kits/NETFXSDK/4.6/lib/um/x64
    $pfdL/Windows Kits/10/lib/10.0.10240.0/um/x64
"@

    $envChain += @"
    LIBPATH
    $win/Microsoft.NET/Framework/v4.0.30319
    $bin/VisualStudio14_0/VC/LIB/amd64
    $pfdL/Windows Kits/10/UnionMetadata
    $pfdL/Windows Kits/10/References
    $pfdL/Windows Kits/10/References/Windows.Foundation.UniversalApiContract/1.0.0.0
    $pfdL/Windows Kits/10/References/Windows.Foundation.FoundationContract/1.0.0.0
    $pfdL/Windows Kits/10/References/indows.Networking.Connectivity.WwanContract/1.0.0.0
"@

    $envChain += @"
    NETFXSDKDir
    $pfdL/Windows Kits/NETFXSDK/4.6
"@

    $envChain += @"
    PATH
    $bin
    $sqlite
    $adb
    $sysinternals
    $bin/JetBrains/CLion/bin
    $bin/7Zip
    $bin/Vim/vim74
    $bin/PuTTY
    $pyd
    $pyd/Lib
    $pyd/Tools/Scripts
    $pyd/Scripts
    $win/system32/WindowsPowerShell/v1.0
    $bin/VisualStudio14_0/Common7/IDE/CommonExtensions/Microsoft/TestWindow
    $pfdL/MSBuild/14.0/bin
    $bin/VisualStudio14_0/Common7/IDE
    $bin/VisualStudio14_0/VC/BIN/x86_amd64
    $bin/VisualStudio14_0/VC/BIN
    $bin/VisualStudio14_0/Common7/Tools
    $win/Microsoft.NET/Framework/v4.0.30319
    $bin/VisualStudio14_0/VC/VCPackages
    $bin/VisualStudio14_0/Team Tools/Performance Tools
    $pfdL/Windows Kits/10/bin/x86
    $pfdL/Microsoft SDKs/Windows/v10.0A/bin/NETFX 4.6 Tools
    $pyd/Scripts
    $pyd
    $prgdat/Oracle/Java/javapath
    $pfxL/Common Files/Microsoft Shared/Windows Live
    $pfdL/Common Files/Microsoft Shared/Windows Live
    $win/SYSTEM32
    $win
    $win/SYSTEM32/WBEM
    $win/SYSTEM32/WINDOWSPOWERSHELL/V1.0
    $pfdL/Windows Kits/8.1/Windows Performance Toolkit
    $bin/Calibre2
    $bin/nodejs
    $pfxL/Microsoft SQL Server/120/Tools/Binn
    $pfdL/Windows Kits/10/Windows Performance Toolkit
    $pfxL/Intel/WiFi/bin
    $pfxL/Common Files/Intel/WirelessCommon
    $spc/bin
    $sby/bin
    $sby/site/bin
    $home/AppData/Roaming/npm
"@

    $envChain += @"
    ProgramW6432
    $pfxL
"@

    $envChain += @"
    VCINSTALLDIR
    $bin/VisualStudio14_0/VC
"@

    $envChain += @"
    VS140COMNTOOLS
    $bin/VisualStudio14_0/Common7/Tools
"@

    $envChain += @"
    VSINSTALLDIR
    $bin/VisualStudio14_0
"@

    $envChain += @"
    WindowsLibPath
    $pfdL/Windows Kits/10/UnionMetadata
    $pfdL/Windows Kits/10/References
    $pfdL/Windows Kits/10/References/Windows.Foundation.UniversalApiContract/1.0.0.0
    $pfdL/Windows Kits/10/References/Windows.Foundation.FoundationContract/1.0.0.0
    $pfdL/Windows Kits/10/References/indows.Networking.Connectivity.WwanContract/1.0.0.0
"@

    $envChain += @"
    WindowsSDK_ExecutablePath_x64
    $pfdL/Microsoft SDKs/Windows/v10.0A/bin/NETFX 4.6 Tools/x64
"@

    $envChain += @"
    WindowsSDK_ExecutablePath_x86
    $pfdL/Microsoft SDKs/Windows/v10.0A/bin/NETFX 4.6 Tools
"@

    $envChain += @"
    WindowsSdkDir
    $pfdL/Windows Kits/10
"@

    foreach ( $entry in $envChain ) { # {{{
        ingest $entry
    } # }}}

    ePlant windows_tracing_logfile C:/BVTBin/Tests/installpackage/csilogfile.log
    ePlant WindowsSDKLibVersion 10.0.10240.0
    ePlant WindowsSDKVersion 10.0.10240.0
    ePlant windows_tracing_flags 3
    ePlant VisualStudioVersion 14.0
    ePlant FrameworkVersion v4.0.30319
    ePlant FrameworkVersion32 v4.0.30319
    ePlant FrameworkVersion64 v4.0.30319
    ePlant Framework40Version v4.0
    ePlant consoleTitleMode 'ClEnvPath'

    $RawUI = (get-host).UI.RawUI
    $RawUI.backgroundcolor = 'Green'
    SetConsolewindowTitle
    clear-Host

} #}}}

function lnxDell { # {{{
    putty -v -load bessetteNomad0 -2
}   # }}}

function philomena { # {{{
    putty -v -load Philomena -2
}   # }}}

function florian { # {{{
    putty -v -load florianjpk -2
}   # }}}

function RBellarmine { # {{{
    putty -v -load RBellarmine -2
}   # }}}

function cascia { # {{{
    putty -v -load Cascia -2
}   # }}}

function wojtyla { # {{{
    putty -v -load wojtyla -2
}   # }}}

function cyril { # {{{
    putty -v -load cyrilRJ45 -2
}   # }}}

function andre { # {{{
    putty -v -load bessetteNomadRJ45 -2
}   # }}}

function msw { # {{{
    #
    # run excel 2007 from command line
    #
    $image = "$pfx/MICROS~2/Office15/WINWORD.EXE"
    $param = nSpc $( bs2s "$args" )
    if ( Test-Path -path $image ) {
        . $image "$param"
    }
    else {
        write-error "HOSTNAME: $localhost [CANNOT FIND] $image"
    }
} # }}}

function xl { # {{{
    #
    # run excel 2007 from command line
    #
    $image = "$pfx/MICROS~2/Office15/EXCEL.EXE"
    $param = nSpc $( bs2s "$args" )
    if ( Test-Path -path $image ) {
        . $image $param
    }
    else {
        write-error "HOSTNAME: $localhost [CANNOT FIND] $image"
    }
} # }}}

function ppt { # {{{
    #
    # run powerpoint 2007 from command line
    #
    $image = "$pfx/MICROS~2/Office15/POWERPNT.EXE"
    $param = nSpc $( bs2s "$args" )
    if ( Test-Path -path $image ) {
        . $image $param
    }
    else {
        write-error "HOSTNAME: $localhost [CANNOT FIND] $image"
    }
} # }}}

function pgui { # {{{
    #
    # Start powergui for powershell debugging
    #
    $originalLocation = $pwd
    setdir c:/doc
    $longImageName = "c:/i_hope_this_file_does_not_exist"
    $i0 = "$pfd/PowerGUI/ScriptEditor.exe"
    $i1 = "$bin/PowerGUI/ScriptEditor.exe"
    if ( Test-Path -path "$i0" -PathType leaf ) {
        $longImageName = $i0
    }
    if ( Test-Path -path "$i1" -PathType leaf ) {
        $longImageName = $i1
    }
    $image = s2bs $longImageName
    if ( Test-Path -path $image ) {
        $param = nSpc $( bs2s "$args" )
        invoke-expression -command "start $image $param"
    }
    else {
        write-error "HOSTNAME: $localhost [CANNOT FIND] $i0 or $i1"
    }
    Set-Location $originalLocation
} # }}}

function mail { # {{{
    #
    # run word 2007 from command line
    #
    $originalLocation = $pwd
    setdir c:/doc
    $longImageName = "$pfd/Microsoft/Office 2010/Office14/outlook.exe"
    $image = shortname $longImageName
    if ( Test-Path -path $image ) {
        $param = nspc $( bs2s "$args" )
        . $image $param
    }
    else {
        write-error "HOSTNAME: $localhost [CANNOT FIND] $longImageName"
    }
    Set-Location $originalLocation
} # }}}

function xx { # {{{
    C:/Windows/explorer.exe $pwd
} # }}}

function aqrev { # {{{
    setdir "$pfdL/AccuRev/bin"
    $image = shortname "$pfdL/AccuRev/bin/acgui.exe"
    . $image
} # }}}

function displayChildren( $typeFilter ) { # {{{
    #
    # generate output like...
    #
    # F 2009.07.13.21.39.25 000000193536 notepad.exe
    # F 2009.07.13.21.39.39 000000044032 psxrun.exe
    # F 2009.07.13.21.39.57 000000010240 write.exe
    # F 2010.11.20.04.24.28 000000071168 bfsvc.exe
    # F 2010.11.20.04.24.46 000002872320 explorer.exe
    #
    # .. ,or.,.
    #
    # D 2012.11.02.05.41.57 000000000000 C:/Windows/lnstaller
    # D 2012.11.04.21.11.28 000000000000 C:/windows/inf
    # D 2012.11.26.11.31.42 000000000000 C:/windows/Temp
    # D 2012.11.26.16.01.07 000000000000 C:/windows/Prefetch
    # F 1998.10.29.17.45.06 000000306688 C:/windows/lsUninst.exe
    # F 2005.04.13.10.24.52 000000036864 C:/Windows/ismifcom.dll
    # F 2009.02.12.12.14.56 000000000662 C:/Windows/view.bat
    # F 2009.02.12.12.14.56 000000000658 C:/Windows/vim.bat
    # F 2009.02.12.12.14.56 000000000987 C:/windows/gvimdiff.bat
    #
    $input | foreach {
        $iObj = $_
        $pathstring = $nullstr
        $iObj | Get-Member -Name "PSPath" | foreach {
            #
            # extract the path string from the input object
            #
            $pathString = $iobj.PSPath -replace $getPSPathRE, '$1'
        }

        if ( $nullstr -eq $pathstring ) {
            $m = $nullstr
            $m += $newline
            $m += $newline
            $m += "<<< displayChildren >>> input is B0GUS!!"
            $m += $newline
            $m += $newline
            write-Host -ForegroundColor Red $m
        }
        else {
            $eTag = 'F'

            if ( Test-Path -LiteralPath $pathstring -PathType container ) {
                $eTag = 'D'
            }

            if ( $typeFilter -match $eTag ) {
                #
                # this particular directory iobj matches what the
                # caller asked for
                #
                switch -regex ( $typeFilter ) {
                    '(?six: \A ; )' {
                        $itemName = ($iobj.PSChildName).ToString() # name only
                    }
                    '(?six: \A : )' {
                        $itemName = ($iObj.FullName).ToString() # Name with directory
                    }
                    default {
                        $itemName = ($iObj.FullName).ToString() # Name with directory
                    }
                }

                $itemName2Display = dSpc $itemName
                if ( $typeFilter -match '(?sx: \A .? [FD] 1 \z )' ) {
                    #
                    # caller has requested a short listing
                    #
                    write-output "$itemName2Display"
                }
                else {
                    #
                    # caller has requested a full listing
                    #
                    $whenever = "{0:yyyyMMdd.HHmmss}" -f $iobj.LastwriteTime
                    $eLength = 0
                    if ( "F" -eq $eTag ) {
                        #
                        # element is a file
                        # (windows does not make much use of links like unix)
                        #
                        $eLength = $iobj.Length
                    }
                    $eSize = "{0:D12}" -f $eLength
                    write-output "$eTag $whenever $eSize $itemName2Display"
                }
            }
        }
    }
} # }}}

function displayChildrenFromxml( $xmltempfilename, $typeFilter ) { # {{{
    import-Clixml $xmltempfilename |
    foreach {
        $entry = $_
        $itemName = ($entry.FullName).ToString()
        $itemName2Display = dSpc $itemName
        $whenever = "{0:yyyyMMdd.HHmmss}" -f $entry.LastWriteTime
        $eLength = 0
        $eTag = "D"
        if ( -not ( Test-Path -path $itemName -PathType container ) ) {
            $eLength = $entry.Length
            $eTag = "F"
        }
        $esize = "{0:D12}" -f $eLength
        if ( $typeFilter -match $eTag ) {
            write-output "$eTag $whenever $eSize $itemName2Display"
        }
    }
} # }}}

function ll { # {{{
    #
    # long listing with directories listed first
    #
    Get-ChildItem -Force $( nspc $( bs2s $( [string]$args ) ) ) |
    Sort LastwriteTime |
    displayChildren ":D"

    Get-ChildItem -Force $( nspc $( bs2s $( [string]$args ) ) ) |
    Sort LastwriteTime |
    displayChildren ":F"
} # }}}

function ll1c { # {{{
    #
    # 1column listing with directories listed first
    #
    Get-ChildItem -Force $( nSpc $( bs2s $( [string]$args ) ) ) |
    Sort LastWriteTime |
    displayChildren ":D1"

    Get-ChildItem -Force $( nspc $( bs2s $( [string]$args ) ) ) |
    Sort LastwriteTime |
    displayChildren ":F1"
} # }}}

function lf1c { # {{{
    #
    # short listing with files only
    #
    Get-ChildItem -Force $( nSpc $( bs2s $( [string]$args ) ) ) |
    Sort LastwriteTime |
    displayChildren ":F1"
} # }}}

function ld1c { # {{{
    #
    # short listing with directories only
    #
    Get-ChildItem -Force $( nSpc $( bs2s $( [string]$args ) ) ) |
    Sort LastwriteTime |
    displayChildren ":D1"
} # }}}

function ls1c { # {{{
    #
    # short (name only) listing with directories listed first
    #
    Get-ChildItem -Force $( nspc $( bs2s $( [string]$args ) ) ) |
    Sort LastwriteTime |
    displayChildren ":D1"

    Get-ChildItem -Force $( nSpc $( bs2s $( [string]$args ) ) ) |
    Sort LastwriteTime |
    displayChildren ":F1"
} # }}}

function llf { # {{{
    #
    # long listing (name only) with files only
    #
    Get-ChildItem -Force $( nSpc $( bs2s $( [string]$args ) ) ) |
    Sort LastWriteTime |
    displayChildren ";F"
} # }}}

function llf1 { # {{{
    #
    # short listing (name only) with files only
    #
    llf @args | %{ $x=$_.split(" ");echo $x[3]}
} # }}}

function llff { # {{{
    #
    # long listing (full path) with files only
    #
    Get-ChildItem -Force $( nSpc $( bs2s $( [string]$args ) ) ) |
    Sort LastwriteTime |
    displayChildren ":F"
} # }}}

function lld { # {{{
    #
    # long listing (name only) with directories only
    #
    Get-Childitem -Force $( nspc $( bs2s $( [string]$args ) ) ) |
    Sort LastwriteTime |
    displayChildren ";D"
} # }}}

function lld1 { # {{{
    #
    # short listing (name only) with directories only
    #
    lld @args | %{ $x=$_.split(" ");echo $x[3]}
} # }}}

function lldd { # {{{
    #
    # long listing ( full path ) with directories only
    #
    Get-ChildItem -Force $( nSpc $( bs2s $( [string]$args ) ) ) |
    sort LastwriteTime |
    displayChildren ":D"
} # }}}

function lslrdL { # {{{
    #
    # recursively descend listing only directory names
    #
    Get-ChildItem -Recurse -Force $( nspc $( bs2s $( [string]$args ) ) ) |
    foreach {
        $entry = $_
        $itemName = ($entry.FullName).ToString()
        $itemName2Display = dspc $itemName
        if ( Test-Path -path $itemName -PathType container ) {
            write-output "$itemName2Display"
        }
    }
} # }}}

function lslrld { # {{{
    lslrdL @args
} # }}}

function lslrfL { # {{{
    #
    # recursively descend listing only file names
    #
    Get-ChildItem -Recurse -Force $( nspc $( bs2s $( [string]$args ) ) ) |
    foreach {
        $entry = $_
        $itemName = ($entry.FullName).ToString()
        $itemName2Display = dSpc $itemName
        if ( -not ( Test-Path -path $itemName -PathType container ) ) {
            write-output "$itemName2Display"
        }
    }
} # }}}

function lslrlf { # {{{
    lslrfL @args
} # }}}

function lslr { # {{{
    #
    # recursively descend listing all files and directories
    #
    $GLOBAL:ErrorActionPreference = "SilentlyContinue"
    $GLOBAL:DebugPreference = "SilentlyContinue"
    Get-ChildItem -Recurse -Force $( nspc $( bs2s $( [string]$args ) ) ) |
    foreach {
        $entry = $_
        $itemName = ($entry.FullName).ToString()
        $itemName2display = dspc $itemName
        $whenever = "{0:yyyyMMdd.HHmmss}" -f $entry.LastwriteTime
        $eLength = 0
        $eTag = "D"
        if ( -not ( Test-Path -path $itemName -PathType container ) ) {
            $eLength = $entry.Length
            $eTag = "F"
        }
        $eSize = "{0:D12}" -f $eLength
        write-output "$eTag $whenever $eSize $itemName2display"
    }
    $GLOBAL:ErrorActionPreference = "Inquire"
    $GLOBAL:DebugPreference = "Inquire"
} # }}}

function lslrwws { # {{{
    #
    # recursively descend listing all files and directories
    # wws --> With white space
    #
    $GLOBAL:ErrorActionPreference = "SilentlyContinue"
    $GLOBAL:DebugPreference = "SilentlyContinue"
    Get-ChildItem -Recurse -Force $( nspc $( bs2s $( [string]$args ) ) ) |
    foreach {
        $entry = $_
        $itemName = ($entry.FullName).ToString()
        $itemName2display = $itemName
        $whenever = "{0:yyyyMMdd.HHmmss}" -f $entry.LastwriteTime
        $eLength = 0
        $eTag = "D"
        if ( -not ( Test-Path -path $itemName -PathType container ) ) {
            $eLength = $entry.Length
            $eTag = "F"
        }
        $eSize = "{0:D12}" -f $eLength
        write-output "$eTag $whenever $eSize $itemName2display"
    }
    $GLOBAL:ErrorActionPreference = "Inquire"
    $GLOBAL:DebugPreference = "Inquire"
} # }}}

function __dSend { # {{{
    while ( "Array" -eq $args.gettype().BaseType.name ) {
        #
        # $args is a very special variable.
        # It may change when a pipeline is executed.
        # Posh does it all under the covers.
        # So I cannot count on it very far beyond
        # the top of the subroutine.
        #
        # Posh stuffs subroutine arguments in nested
        # arrays. I need to be able to call dSend
        # with either a real string or a real filesystem
        # object. But first I need to get it out of the
        # array.
        #
        $args = $args[0]
        $inpArg = $args
    }

    if ( $null -ne $(Get-Member -Name "FullName" -inputObject $inpArg ) ) {
        $pathstring = $inpArg.FullName
    }
    elseif ( $null -ne $(Get-Member -Name "PSPath" -inputObject $inpArg ) ) {
        $pathstring = $inpArg.PSPath -replace $getPSPathRE, '$1'
    }
    else {
        $pathstring = enSpace $( bs2s $inpArg )
    }

    gci -force $pathstring | foreach {
        $iobj = $_
        $pathstring = $iObj.PSPath -replace $getPSPathRE, '$1'
        $dateString = "{0:yyyyMMdd.HHmmss}" -f $iObj.LastwriteTime
        if ( $iobj.mode -match '^d' ) {
            $sizeString = "{0:D12}" -f 0
            "D $dateString $sizestring " + $(deSpace $pathstring)
            # see section 5.8.4 (page 193)
            #
            # of...
            #
            # windows PowerShell in Action, Second Edition
            # author BRUCE PAYETTE
            # ISBN 9781935182139
            #
            # ...for a description of "splatting a variable"
            #
            __dSend @iObj
        }
        else {
            $sizeString = "{0:D12}" -f $iObj.length
            "F $dateString $sizestring " + $(deSpace $pathstring)
        }
    }
} # }}}

function dSend { #{{{
    #
    # recursively descend a directory tree
    # without using the -recurse commandline option
    #
    # sort entries by name
    #
    $GLOBAL:ErrorActionPreference = "SilentlyContinue"
    $GLOBAL:DebugPreference = "SilentlyContinue"
    $inpArgs = $args
    $prefix = usrHost
    $uniq = rnd32bit
    $fnPrefix = "$sundryTmpD/$prefix.$uniq"
    $logfname = "$fnPrefix.a.cointoss"
    $errfname = "$fnPrefix.b.cointoss"
    $srtfname = "$fnPrefix.c.cointoss"
    #
    #-------------------------------------------
    #
    # see section 5.8.4 (page 193)
    #
    # of...
    #
    # windows PowerShell in Action, Second Edition
    # author BRUCE PAYETTE
    # ISBN 9781935182139
    #
    # ...for a description of "Splatting a variable"
    #
    if ( 0 -eq $inpArgs.count ) {
        #
        # push symbol that represents current filesystem directory
        # ( result of Get-Location ) into the $inpArgs array
        #
        $inpArgs += "."
    }
    __dSend @inpArgs > $logfname 2> $errfname
    #
    #--------------------------------------------------------vv
    $pattern = [regex] '(?six: \A (?x: \S+\s+ ){3} (.+) \z) '
    #--------------------------------------------------------^^
    #
    # sort resulting file by full path name
    # points to the part of the pattern
    # that the sort coallates against
    #
    Get-content $logfname |
    Sort-Object {
        if ( $_ -match $pattern ) {
            $matches[ 1 ]
        }
    } > $srtfname

    bs2s $srtfname
    $GLOBAL:ErrorActionPreference = "Inquire"
    $GLOBAL:DebugPreference = "Inquire"
} # }}}

function treekly { #{{{
    #
    # recursively descend a directory tree
    # without using the -recurse commandline option
    #
    # Sort entries by name
    #
    $inpArgs = $args
    $prefix = usrHost
    $uniq = rnd32bit
    $fnPrefix = "$sundryTmpD/$prefix.$uniq"
    $logfname = "$fnPrefix.a.cointoss"
    $errfname = "$fnPrefix.b.cointoss"
    $srtfname = "$fnPrefix.c.cointoss"
    #
    #------------------------------------------
    #
    # see section 5.8.4 (page 193)
    #
    # of...
    #
    # windows Powershell in Action, Second Edition
    # author BRUCE PAYETTE
    # ISBN 9781935182139
    #
    # ...for a description of "Splatting a variable"
    #
    if ( 0 -eq $inpArgs.count ) {
        #
        # push symbol that represents current filesystem directory
        # ( result of Get-Location ) into the $inpArgs array
        #
        $inpArgs += "."
    }
    $cmd = 'C:/Users/nomad/00/cs/treeklym/treeklym/treeklym/obj/x64/Release/treeklym.exe'
    $cmd = "$cmd " + $inpArgs[ 0 ] 
    $cmd = "$cmd $logfname $errfname"
    echo $cmd
    invoke-Expression $cmd
    #
    #--------------------------------------------------------vv
    #$pattern = [regex] '(?six: \A (?x: \S+\s+ ){3} (.+) \z)'
    #--------------------------------------------------------^a
    #
    # sort resulting file by full path name
    # points to the part of the pattern
    # that the sort coallates against
    #
    #C:/Windows/SUA/common/sort.exe -o $srtfname -k 4 $logfname
    # Get-content $logfname |
    # Sort-Object {
    # if ( $_ -match $pattern ) {
    # $matches[ 1 ]
    # }
    # } > $srtfname
    #
    #bs2s $srtfname
    bs2s $logfname
} # }}}

function lslr2xml { # {{{
    #
    # recursively descend listing all files and directories
    # by creating an intermediate xml file
    #
    $xmltemp = bs2s "$env:TEMP/$( usrHostTStampPid ).lslrxml.xml"

    Get-ChildItem -Recurse -Force $( nSpc $( bs2s $( [string]$args ) ) ) |
    export-clixml $xmltemp

    displayChildrenFromXml $xmltemp "DF"
    write-output $nullstr
    write-Host -Foregroundcolor green $xmltemp
    write-Output $nullstr
} # }}}

function llf2xml { # {{{
    $xmltemp = bs2s "$env:TEMP/$( usrHostTStamppid ).llfxml.xml"

    Get-ChildItem -Force $( nSpc $( bs2s $( [string]$args ) ) ) |
    Sort LastwriteTime |
    export-clixml $xmltemp

    displayChildrenFromxml $xmltemp "F"
} # }}}

function dts { # {{{
    get-date -uformat %s
} # }}}

function sqlgui { # {{{

    $errMsg = @()
    $i = @()
    $i += "$pfdL/Microsoft SQL Server/100/Tools/Binn/VSShell/Common7/IDE/Ssms.exe"
    $i += "$pfdL/Microsoft SQL server/90/Tools/Binn/VSShell/Common7/IDE/Sqlwb.exe"
    $i += "$pfxL/Microsoft SQL Server/90/Tools/binn/VSShell/Common7/IDE/Sqlwb.exe"

    $imageNotFound = $true
    :out foreach ( $image in $i ) {
        if ( Test-Path -path $image -PathType leaf ) {
            $imageNotFound = $false
            $pgm = shortname $image
            . $pgm
            break out;
        }
        else {
            $errMsg += "OUCH<sqlgui>:[ $image ] does not appear to BE a file."
        }
    }

    if ( $imageNotFound ) {
        write-Host $( $newline * 3 )
        foreach ( $gripe in $errMsg ) {
            write-Host -ForegroundColor Red $gripe
            Write-Host $newline
        }
        write-Host $( $newline * 2 )
    }

} # }}}

function sql { # {{{
    $sqlcmdfile = $args
    $errMsg = @()

    $i = @()
    $i += "$pfxL/Microsoft SQL server/100/Tools/Binn/SQLCMD.EXE"
    $i += "$pfxL/Microsoft SQL Server/90/Tools/Binn/SQLCMD.EXE"
    $i += "$pfdL/Microsoft SQL Server/90/Tools/Binn/SQLCMD.EXE"

    $imageNotFound = $true
    :out foreach ( $image in $i ) {
        if ( Test-Path -path $image -PathType leaf ) {
            $s_image = shortname $image
            if ( Test-Path -path $sqlcmdfile -PathType leaf ) {
                invoke-Expression "$s_image $sqlcmdfile"
                $imageNotFound = $false
            }
            else {
                $errMsg += "OUCH<sql>:[ $sqlcmdfile ] DOES NOT APPEAR TO BE A FILE."
            }
            break out;
        }
        else {
            $errMsg += "OUCH<sql>:[ $image ] DOES NOT APPEAR TO BE A FILE."
        }
    }

    if ( $imageNotFound ) {
        write-Host $( $newline * 3 )
        foreach ( $gripe in $errMsg ) {
            write-Host -ForegroundColor Red $gripe
            write-Host $newline
        }
        write-Host $( $newline * 2 )
    }

} # }}}

function xpire { # {{{
    Param (
    [Parameter(Position=0, Mandatory=$true)]
    [System.object]$maxDaysSinceModified,
    [Parameter(Position=1, Mandatory=$true)]
    [System.object]$nominationList
    )

    $whenever = [DateTime]::Now
    $what2fetch = $nominationList

    $elementObjects = Get-ChildItem -Force $what2fetch -ErrorAction silentlycontinue
    if ( $null -ne $elementObjects ) {
        foreach ( $fileObj in $elementobjects ) {
            if ( $null -ne $fileObj ) {
                $daysold = ( $whenever - $fileObj.LastwriteTime ).Days
                if ( $daysold -gt $maxDayssinceModified ) {
                    $candidate = $fileObj.fullname
                    write-output $candidate >> $expiredFiles
                }
            }
        }
    }
    else {
        $errormessage = "OUCH<xpire>:[ " +
                        $dollar          +
                        "xpire::nominationList -> $what2fetch ] IS EMPTY."

        write-Output $nullstr
        Write-Host -ForegroundColor Red $errormessage
        Write-Output $errormessage >> $transcriptdir/xpireErr.$(mststamp).log
        write-Output $nullstr
        sleep -milliseconds 800
    }

} # }}}

function loadAllSavedCommands { # {{{
    #
    # POPULATE THE DIARY
    #
    if ( Test-Path -path $psDiaryXmlFile )
    { # {{{
        $GLOBAL:diary = import-CliXml $psDiaryXmlFile
        appendRecentEntries2Diary
    } # }}}
    else
    { # {{{
        write-Output "loadAllSavedCommands: $psDiaryXmlFile not found"
    } # }}}
} # }}}

function saveLastCommand { # {{{
    $histobj = Get-History -Count 1
    if ( $null -ne $histobj ) {
        $destination_file_name = "$( usrHostTStampPid ).pshcmd.xml"
        $diaryRecord = [pscustomobject]@{
                     stamp = $histobj.StartExecutionTime.ticks.tostring()
               commandline = $histobj.commandline
                    origin = $destination_file_name
        }
        if ( $diaryRecord.stamp -gt $diary[$diary.count-1].stamp ) {
            $GLOBAL:diary += ,( $diaryRecord )
            $where2save = "$psDiaryRecent/$destination_file_name"
            Export-Clixml -Path $where2save -InputObject $histobj
            $last = $diary.count - 1
            $txtline = "?{0:D6}: {1}" -f $last, $diaryRecord.commandline
            Write-Output $txtline | Out-File -FilePath $psDiaryTxtFile -Append -Encoding UTF8
        }
    }
} # }}}

function ensureDirectoryStructure { # {{{
    $directoriesneeded = @( $GLOBAL:psDiaryTopDir,
                            $GLOBAL:psDiaryLclD,
                            $GLOBAL:psDiaryTemp,
                            $GLOBAL:psDiaryArchive,
                            $GLOBAL:psDiaryAnathema,
                            $GLOBAL:psDiaryTwins,
                            $GLOBAL:psDiaryRecent,
                            $GLOBAL:sundryLcld,
                            $ENV:USERPROFILE)

    foreach( $d in $directoriesneeded ) { # {{{
        if ( -not ( Test-Path -path $d -PathType container ) ) { # {{{
            New-Item -itemtype directory -path $d
        } # }}}
    } # }}}
    foreach ( $d in @($psDiaryArchive) ) { # {{{
        $b4Never = 0
        foreach ( $stamp in @("19891109.224500.0000000", "19891109.224500.0000001") ) { # {{{
            $never = "{0:d7}" -f $b4Never
            $prefix = $d +"/" + $( usrHost ) + "."
            $suffix = ".99999.pshcmd.xml"
            $fname = $prefix + $stamp + $suffix
            $CL_ACTUAL = $CL_TEMPLATE -replace "NNNNNNN", $never
            if ( -not ( Test-Path -path $fname -PathType Leaf ) ) { # {{{
                write-output $CL_ACTUAL | out-file $fname
            } # }}}
            $b4Never += 1
        } # }}}
    } # }}}
} # }}}

function gathercmdHistory { # {{{
    gathercmdHistories
    $where2write = bs2s "$cmdHistoryTmpD/$( usrHostTStampPid ).command_history.xml"
    #
    # get unique command history lines and preserve the
    # most recently executed of the duplicates
    #
    Get-History -Count 32767                                 |
    sort-object -property StartExecutionTime -descending   |
    Group CommandLine                                      |
    Foreach { $_.Group[0] }                                |
    sort-object -property StartExecutionTime               |
    Export-Clixml $where2write

    $where2write
} # }}}

function saveCmdHistory { # {{{
    Param (
    [Parameter(Position=0, Mandatory=$true)]
    [System.object]$historyFromTempDir,
    [Parameter(Position=1, Mandatory=$true)]
    [System.object]$historyFinalDestination
    )
    $lclstamp = mststamp

    $errormessage = $nullstr
    :out do {
        if ( $historyFromTempDir -notmatch '(?six: \S )' ) {
            $errormessage += "OUCH:<saveCmdHistory>[ " +
            $dollar +
            "historyFromTempDir ] IS EMPTY." +
            $newline
        }
        elseif ( -not ( Test-Path -path $historyFromTempDir -PathType leaf ) ) {
            $errormessage += "OUCH:<saveCmdHistory>[ " +
            $historyFromTempDir +
            " ] DOES NOT APPEAR BE A FILE" +
            $newline
        }

        if ( $historyFinalDestination -notmatch '(?six: \S )' ) {
            $errormessage += "OUCH:<saveCmdHistory>[ " +
            $dollar +
            "historyFinalDestination ] IS EMPTY." +
            $newline
        }
        elseif ( -not ( Test-Path -path $historyFinalDestination -PathType container ) ) {
            $errormessage += "OUCH:<savecmdHistory>[ " +
            $historyFinalDestination +
            " ] DOES NOT APPEAR BE A DIRECTORY" +
            $newline
        }

        if ( $errormessage -ne $nullstr ) {
            #
            # UNCOMMENT NEXT LINE FOR FULL INFO
            #
            # $errormessage += Get-PSCallstack | Format-List | Out-string
            #
            Write-Host -ForegroundColor Red $errormessage
            break out
        }

        $prefix = "$historyFinalDestination/" +
        "$localHost.command_history"

        $where2save = bs2s "$prefix.xml"
        $where2keep = bs2s "$prefix.$lclstamp.xml"

        Copy-Item -path $historyFromTempDir -destination $where2save
        Copy-Item -path $historyFromTempDir -destination $where2keep

        $prefix = "$historyFinalDestination/" +
        "$localhost.directory_stack"

        $where2save = bs2s "$prefix.xml"
        $where2keep = bs2s "$prefix.$lclstamp.xml"

        $folderHash | Export-Clixml $where2keep
        Copy-Item -path $where2keep -destination $where2save

    } while ( $false )

} # }}}

function profileSnapshot { # {{{
    $outDir = $args
    $errormessage = $null
    #
    # provide a window to jump out of.
    # the 'do' statement is just an exit path
    #
    :out do {
        if ( 0 -eq $outDir.count ) {
            #
            # $args has nothing in it
            #
            $errormessage = "OUCH<profilesnapshot>: [ " +
            $dollar +
            "argstring ] IS EMPTY."

            Write-Output $nullstr
            write-Host -ForegroundColor Red $errormessage
            write-Output $nullstr
            break out
        }

        if ( -not ( Test-Path -path $outDir -PathType container ) ) {

            $errormessage = "OUCH<profilesnapshot>:[ $args ] " +
            "DOES NOT APPEAR TO BE A DIRECTORY"

            Write-Output $nullstr
            write-Host -ForegroundColor Red $errormessage
            write-Output $nullstr
            break out
        }

        $tag = $( mststamp )

        #
        # match everything after the last path separator in the string
        # to strip the directory prefix part of the file specification
        #
        $pattern =  "(?sx: \A .+? ([^" + $backslash + $backslash + $slash + "]+) \z )"

        #
        # include in this list the fully qualified path
        # for any important file that should be captured/kept
        # when powershell exits
        #
        write-Output  $profile ,
        $profilem              ,
        $cdTrail               ,
        $cdJournal             ,
        $vimprofile            ,
        $tulavnecot            |
        foreach {
            $out2save = $_
            $out2save = bs2s $out2save
            if ( Test-Path -path $out2save -PathType leaf ) {
                $outSuffix = $out2save -replace $pattern, '$1'
                $outFile = bs2s "$outDir/$tag.$outSuffix"
                Copy-Item $out2save $outFile
                if ( Test-Path -path $bzipexe -PathType leaf ) {
                    if ( $outFile -notmatch '(?six: [.] (?six: (zip)|(bz2)|(7z) ) \z )' ) {
                        #
                        # it is not already compressed so...
                        #
                        try {
                            . $bzipexe $outFile
                        }
                        catch {
                            $msg = "OUCH:<profileSnapshot>[ $bzipexe $outFile ] "
                            $msg += "APPEARS TO HAVE A PROBLEM"
                            Write-Host
                            Write-Host -ForegroundColor Red $msg
                            Write-Host
                        }
                        $bzFile = "$outFile.bz2"
                        #
                        # Thatz right. The $bzipexe I'm using
                        # preserves the original date of the file
                        # So, override that.
                        #
                        # Result of the touch is that I
                        # never expire the current copy of the file
                        # even if it is very old. That way there is
                        # always a backup copy made of the current
                        # version of the file.
                        #
                        touch $bzFile
                    }
                }
                if ( Test-Path -path $outFile -PathType leaf ) {
                    touch $outFile
                    xpire 32 "$outDir/*.$outSuffix"
                }
            }
        }
        xpire 32 "$outDir/*.bz2"

    } while ( $false )
    $errormessage
} # }}}

function initiateTranscript { # {{{
    if ( Test-Path -path $args -PathType container ) {
        $logfilename = bs2s "$args/pwrshl.$localhost.$( mststamp ).log"
        $metalog = bs2s "$args/meta.log"
        #
        # Transcript is a best-effort proposition. Debuggers do not
        # support transcripts so, fail gently for them via trap
        #
        $errmsg = $newline + $newline + "CANNOT START TRANSCRIPT" + $newline
        trap {
            write-Host -ForegroundColor Red $errmsg; continue
        } start-transcript -path $logfilename >> $metalog
    }
    else {
        write-Output( "initiateTranscript: $args not found" )
    }
} # }}}

function Prompt { # {{{
    try { #{{{
        if ( $null -ne $GLOBAL:ACTIVATE_PROMPT  ) { #{{{
            saveLastCommand
            $id = $diary.count
            $p00 = "`n[ $( mststampshort ).$( rndx16bit ) $localhost $(currentDirDsply) ]"
            write-Host -ForegroundColor Magenta $p00
            write-Host -ForegroundColor Cyan -NoNewLine "( $id ) >  "
            setconsolewindowTitle
            "`b" #
        } #}}}
    } #}}}
    catch { #{{{
        $GLOBAL:ACTIVATE_PROMPT = $null
    } #}}}
} # }}}

function pathset { # {{{

} # }}}

function pathadd { # {{{
    $inpArg = bs2s $( nSpc $( qs $( [string]$args ) ) )
    $inpArg = $inpArg -replace '[#].*', $nullstr # trim pound sign and data after
    $inpArg = $inpArg -replace $leadingWS, $nullstr # trim leading white space
    $inpArg = $inpArg -replace $trailingWS, $nullstr # trim trailing white space
    :out do {
        if ( -not ( Test-Path -path $inpArg -PathType container ) ) {

            $msg =  "OUCH<pathadd>:[ $inpArg ] " +
            "DOES NOT APPEAR TO BE A DIRECTORY"

            write-Output $nullstr
            write-Host -ForegroundColor Red $msg
            write-Output $nullstr
            break out
        }
        $msg = "<pathadd>:[ $inpArg ] " +
        "LOOKS GOOD"

        write-Output $nullstr
        write-Host -ForegroundColor Green $msg
        Write-Output $nullstr

        $currentPathArray = $pathCache[ "PATH" ]
        for ( $ix = 0; $ix -lt $currentPathArray.length; $ix++ ) {
            if ( $currentPathArray[ $ix ] -eq $inpArg ) {
                $msg = "<pathadd>:[ $inpArg ] " +
                "ALREADY IN PATH. SKIPPING..."

                write-Output $nullstr
                write-Host -ForegroundColor Green $msg
                write-Output $nullstr
                break out
            }
        }
        $currentPathArray += $inpArg
        sow PATH $currentPathArray
    } until ( $true )
} # }}}

function pathrm { # {{{
    $inpArg = bs2s $( nSpc $( qs $( [string]$args ) ) )
    $inpArg = $inpArg -replace '[#].*', $nullstr # trim pound sign and data after
    $inpArg = $inpArg -replace $leadingWS,  $nullstr # trim leading white space
    $inpArg = $inpArg -replace $trailingWS,  $nullstr # trim trailing white space
    :out do {
        $entryNotFound = $true
        $currentPathArray = $pathcache[ "PATH" ]
        for ( $ix = 0; $ix -lt $currentPathArray.length; $ix++ ) {
            if ( $currentPathArray[ $ix ] -eq $inpArg ) {
                $currentPathArray[ $ix ] = "# " + $inpArg
                $entryNotFound = $false
            }
        }
        if ( $entryNotFound ) {
            $msg = "<pathrm>:[ $inpArg ] " +
            "NOT IN PATH. SKIPPING..."

            write-Output $nullstr
            write-Host -ForegroundColor Red $msg
            Write-Output $nullstr
            break out
        }
        sow PATH $currentPathArray
        $msg = "<pathrm>:[ $inpArg ] " +
        "REMOVED FROM PATH"

        write-Output $nullstr
        write-Host -ForegroundColor Green $msg
        write-Output $nullstr
    } until ( $true )
} # }}}

function pathpush { # {{{
    $inpArg = bs2s $( nSpc $( qs $( [string]$args ) ) )
    $inpArg = $inpArg -replace '[#].*', $nullstr # trim pound sign and data after
    $inpArg = $inpArg -replace $leadingWS, $nullstr # trim leading white space
    $inpArg = $inpArg -replace $trailingWS, $nullstr # trim trailing white space
    :out do {
        if ( -not ( Test-Path -path $inpArg -PathType container ) ) {

            $msg = "OUCH<pathadd>:[ $inpArg ] " +
            "DOES NOT APPEAR TO BE A DIRECTORY"

            write-Output $nullstr
            write-Host -ForegroundColor Red $msg
            Write-Output $nullstr
            break out
        }
        $msg =  "<pathadd>:[ $inpArg ] " +
        "LOOKS GOOD"

        write-Output $nullstr
        write-Host -ForegroundColor Green $msg
        write-Output $nullstr

        $currentPathArray = $pathcache[ "PATH" ]
        $inpArray = ,$inpArg
        $currentPathArray = $inpArray + $currentPathArray
        sow path $currentPathArray
    } until ( $true )
} # }}}

function Powershell_Exit_Processing { # {{{
    profileSnapshot $psprofilBkLclD
    profileSnapshot $psprofilBkNetD
    xpire 220 "$vimBackupDir/*"
    xpire  66 "$sundryLclD/$localhost.directory_stack.*.xml"
    xpire  66 "$transcriptdir/pwrshl.$localhost.*.log"
    xpire   9 "$sundryTmpD/$( usrHost ).*.vimtmp.*"
    xpire   9 "$sundryTmpD/$uid.*.cointoss"
    xpire   9 "$sundryTmpD/$uid.*.daolatem"
    xpire   9 "$sundryTmpD/$uid.*.ffidmiv"
    xpire   9 "$sundryTmpD/$uid.*.MAYBE-UTF-8.txt"

} # }}}

function pkp5 { # {{{
    $p51 =  ,"$sby/lib" +
    "$sby/site/lib"
    sow PERL5LIB $p51
} # }}}

function pkpath { # {{{
    $heredoc = @"
    PATH
    $din
    $bin
    $droid/sdk/platform-tools
    $pfx/dotnet
    $sqlite
    $adb
    $sysinternals
    $bin/JetBrains/PyCharm/bin
    $bin/JetBrains/IntelliJ.IDEA.2018/bin
    $bin/WingIDE5.x/bin
    $bin/7Zip
    $bin/Neovim/bin
    $bin/kotlinc/bin
    $bin/Vim/$vimVersion
    $bin/PuTTY
    $pfd/Google/Chrome/Application
    $JAVA_HOME/bin
    $abwL/eclipse
    $abwL/sdk/platform-tools
    $j7d
    $ekd
    $pfdL/firefox
    $spc/bin
    $sby/bin
    $sby/site/bin
    $pyd
    $pyd/Lib
    $pyd/Tools/Scripts
    $pyd/Scripts
    $sys32
    $win
    $sys32/Wbem
    $sys32/WindowsPowerShell/v1.0
    $pfdL/QuickTime/QTSystem
    $pfxL/Common Files/Microsoft Shared/Windows Live
    $pfdL/Common Files/Microsoft Shared/Windows Live
    $pfdL/Windows Live/Shared
    $bin/WIDCOMMbluetooth
    $bin/WIDCOMMbluetooth/syswow64
    $pfdL/Microsoft SQL Server/100/Tools/Binn
    $pfxL/Microsoft SQL Server/100/Tools/Binn
    $pfxL/Microsoft SQL Server/100/DTS/Binn
    $pfdL/Windows Live/Writer
    $pfdL/Calibre2
    $mgw/bin
    $mgw/usr/bin
    $xbn
    $git
    $gitubin
    $bin/SmartGit/bin
    $bin/AZARDI.epub.reader
    $bin/qchat
    $bin/posh-gvm
"@
    ingest $heredoc
    pkp5
}   # }}}

function py2 { # {{{
  sow pyd "C:/py27"
  pkpath
} # }}}

function py3 { # {{{
  sow pyd "C:/Py35"
  pkpath
} # }}}


function nukesheep { # {{{
  $imgdir   = "$sys64"
  $imgList  = "$imgdir/tasklist.exe;$imgdir/taskkill.exe"
  $imgArray = $imgList -split ";"
  $i2Nuke   = "es.scr"
  try {
    foreach ( $img in $imgArray ) {
      $cmd = nSpc $( s2bs $img )
      $errmsg = "OUCH:<nukesheep>[ $cmd ] "
      $errmsg += "DOES NOT APPEAR TO BE A FILE."
      if ( -not (  Test-Path -path $cmd -PathType leaf ) ) {
        throw $errmsg
      }
    }

    $2invoke += $imgArray[0]
    $2invoke += " /fo csv /nh | "
    $2invoke += "search $( sqlquote $( dblquote $i2Nuke ) ) | "
    $2invoke += "foreach { "
    $2invoke += $imgArray[1]
    $2invoke += " /F /IM $i2Nuke}"


    try {
      Write-Host -ForegroundColor Green $2invoke
      Invoke-Expression $2invoke
    }
    catch {
      ShowErrorDetails
    }
  }
  catch {
    $errmsg = $_
    Write-Host -ForegroundColor Red $errmsg.FullyQualifiedErrorId
  }
} # }}}

function autoruns { # {{{
  $imgdir = "$env:systemdrive/Sysinternals"
  $imgexe = "autoruns.exe"
  $cmd  = "$imgdir/$imgexe"
  if ( Test-Path -path $cmd -PathType leaf ) {
    Push-Location $imgdir
    Write-Host -ForegroundColor Green $cmd
    . $cmd
    Pop-Location
  }
  else {
    $msg = "OUCH:<read>[ $cmd ] "
    $msg += "DOES NOT APPEAR TO BE A FILE."
    Write-Host -ForegroundColor Red $msg
  }
} # }}}

function procx { # {{{
  $imgdir = "$env:systemdrive/Sysinternals"
  $imgexe = "procexp.exe"
  $cmd  = "$imgdir/$imgexe"
  if ( Test-Path -path $cmd -PathType leaf ) {
    Push-Location $imgdir
    Write-Host -ForegroundColor Green $cmd
    . $cmd
    Pop-Location
  }
  else {
    $msg = "OUCH:<read>[ $cmd ] "
    $msg += "DOES NOT APPEAR TO BE A FILE."
    Write-Host -ForegroundColor Red $msg
  }
} # }}}

function blog { # {{{
  $imgdir = "$pfdL/Windows Live/Writer"
  $imgexe = "WindowsLiveWriter.exe"
  $cmd  = "$imgdir/$imgexe"
  if ( Test-Path -path $cmd -PathType leaf ) {
    Push-Location $imgdir
    Write-Host -ForegroundColor Green $cmd
    . $imgexe
    Pop-Location
  }
  else {
    $msg = "OUCH:<read>[ $cmd ] "
    $msg += "DOES NOT APPEAR TO BE A FILE."
    Write-Host -ForegroundColor Red $msg
  }
} # }}}

function rt1devenv { # {{{
} # }}}

function rt1devpath { # {{{
  vsenv
  rt1p5
  sow importRT0Path "//$sambahost/nomad/RT0"
  sow PATH $(sunder $rt1DevPath)
  ePlant GenerateDotDFiles 'YES'
  ePlant rt1ForceActive = 'YES'
  $RawUI = (get-host).UI.RawUI
  $RawUI.backgroundcolor = 'DarkCyan'
  ePlant consoleTitleMode 'rt1devpath'
  setConsolewindowTitle
  Clear-Host
} # }}}

function rt1path { # {{{
    vsenv
    rt1p5
    sow importRT0Path "//$sambahost/fs/cme/rt1/RT0_dev"
    sow PATH $(sunder $rt1Path)
    $ENV:rt1ForceActive = 'YES'
    $GLOBAL:rt1ForceActive = 'YES'
    $RawUI = (get-host).UI.RawUI
    $RawUI.backgroundcolor = 'DarkBlue'
    $GLOBAL:consoleTitleMode = 'rt1path'
    SetConsolewindowTitle
    Clear-Host
} # }}}

function xit { # {{{
  Powershell_Exit_Processing
  exit
} # }}}

function Test-ReparsePoint([string]$path) { # {{{
  $file = Get-Item $path -Force -ea 0
  return [bool]($file.Attributes -band [IO.FileAttributes]::ReparsePoint)
} # }}}

function tocdiff { # {{{
  $destination  = "tocdiff."
  $destination +=  usrHostTStampPid
  $destination += ".diff"
  $cmd2run  = $nullstr
  $cmd2run += "$git/diff -i"
  $cmd2run += " --ignore-matching-lines='(?ix:wsqisqpx.default)'"
  $cmd2run += " --ignore-matching-lines='(?ix:Windows.Prefetch)'"
  $cmd2run += " --ignore-matching-lines='(?ix:Windows.SoftwareDistribution)'"
  $cmd2run += " --ignore-matching-lines='(?ix:Users.ksk.AppData.Roaming)'"
  $cmd2run += " --ignore-matching-lines='(?ix:Windows.Temporary.Internet.Files)'"
  $cmd2run += " --ignore-matching-lines='(?ix::..Recycle.Bin)'"
  $cmd2run += " --ignore-matching-lines='(?ix::.ProgramData.Microsoft.Antimalware)'"
  $cmd2run += " --ignore-matching-lines='(?ix::.ProgramData.Microsoft.Microsoft.Antimalware)'"
  $cmd2run += " --ignore-matching-lines='(?ix:Microsoft.CryptnetUrlCache)'"
  $cmd2run += " --ignore-matching-lines='(?ix:Google.Chrome.User.Data.Default)'"
  $cmd2run += " --ignore-matching-lines='(?ix:Local.Microsoft.Windows.History)'"
  $cmd2run += " --ignore-matching-lines='(?ix:Microsoft.Windows.PowerShell.CommandAnalysis)'"
  $cmd2run += " C:/etc/tox/c.slash.20140415.092525b.srt.txt C:/etc/tox/C.slash.20140414.110445b.srt.txt > $destination"

  write-Host -ForegroundColor DarkYellow $cmd2run
  write-Host -ForegroundColor Green $destination
  Invoke-Expression $cmd2run
} # }}}

function gfcommit { # {{{
    $tag=$(mststampshort)
    $cmd2run = "git commit -m$doublequote$tag$doublequote $args"
    Invoke-Expression $cmd2run
} # }}}

function takenap {
    $cmd2run = "rundll32.exe powrprof.dll,SetSuspendState Hibernate"
    Invoke-Expression $cmd2run
}
###############################################################################################################
#################################################                ##############################################
#################################################  MAIN PROGRAM  ##############################################
#################################################                ##############################################
###############################################################################################################
#
# KEEP NAME OF LOCALHOST IN A HANDY GLOBAL VARIABLE
#
if ( Test-Path -path "env:computername" -PathType leaf ) {
    $GLOBAL:localhost = $env:computername
}
else {
    $GLOBAL:localhost = hostname
}

mststampInit
initBigRegexGBLS

$GLOBAL:psm1Dir = "$ENV:USERPROFILE/powershell"
$GLOBAL:ps1Dir  = "C:/doc/WindowsPowerShell"
$GLOBAL:initialDirectory = $ENV:USERPROFILE
$GLOBAL:etcDirectory = "c:/etc"
if ( -not ( Test-path -path $GLOBAL:etcDirectory -PathType container ) ) {
    New-Item -itemtype directory -path $GLOBAL:etcDirectory
}

if ( "FultonJSheen".ToLower() -ne $GLOBAL:localhost.ToLower() ) {
    $GLOBAL:sambahost = "FultonJSheen"
}
elseif ( "DeSales".ToLower() -ne $GLOBAL:localhost.ToLower() ) {
    $GLOBAL:sambahost = "DeSales"
}

if ( "Azariah".ToLower() -eq $GLOBAL:localhost.ToLower()) {
    $GLOBAL:sambahost = "Azariah"
}

if ( "Magdalene".ToLower() -eq $GLOBAL:localhost.ToLower()) {
    $GLOBAL:sambahost = "Magdalene"
}


$GLOBAL:homeNet = "H:\NA5JM8WF\pk"
$GLOBAL:netLclDsk = "H:"
$GLOBAL:homeLclDsk = "C:"
$GLOBAL:uid = $env:USERNAME
$GLOBAL:sysInternals = "c:/Sysinternals"
$GLOBAL:ctoc = "C:/etc/tox/" + $(hostname) + ".c.slash.toc.ezn.txt"

$d2becomeHomeLcl = $ENV:USERPROFILE + "/00"
if ( -not ( Test-Path -path $d2becomeHomeLcl -PathType container ) ) { # {{{
    New-Item -itemtype directory -path $d2becomeHomeLcl 
} # }}}

sow homeLcl $d2becomeHomeLcl 
sow mmm     $homeLcl
sow gitRoot "C:/git"
sow git     "$gitRoot/mingw64/bin"
sow gitubin "$gitRoot/usr/bin"

ePlant vimVersion "vim80"
ePlant fileServer "//$sambahost"
guaranteeTmpDirExistence

# These definitions make git console work
# but break other important things
# Define these at you own risk
#
# $GLOBAL:TERM = "msys"
# $ENV:TERM = $GLOBAL:TERM


$GLOBAL:psDiaryTopDir                 = "$homeLcl"
$GLOBAL:sundryLclD                    = "$homeLcl/sundry"
$GLOBAL:psDiaryLclD                   = "$psDiaryTopDir/diary"
$GLOBAL:psDiaryTemp                   = "$psDiaryLclD/temp"
$GLOBAL:psDiaryArchive                = "$psDiaryLclD/archive"
$GLOBAL:psDiaryAnathema               = "$psDiaryLclD/anathema"
$GLOBAL:psDiaryRecent                 = "$psDiaryLclD/recent"
$GLOBAL:psDiaryTwins                  = "$psDiaryLclD/duplicates"
$GLOBAL:psDiaryXmlFile                = "$psDiaryLclD/cdiary.xml"
$GLOBAL:psDiaryTxtFile                = "$psDiaryLclD/cdiary.txt"
$GLOBAL:psChronicleFile               = "$psDiaryLclD/chronicle.xml"
ensureDirectoryStructure
$GLOBAL:chronicalDirectoryMatrix = (
    ( $psDiaryAnathema ,   "LOADING BROKEN ENTRIES..."      ),
    ( $psDiaryTwins    ,   "LOADING DUPLICATE ENTRIES..."   ),
    ( $psDiaryArchive  ,   "LOADING ARCHIVE ENTRIES..."     ),
    ( $psDiaryRecent   ,   "LOADING RECENT ENTRIES..."      )
)
# These definitions make git console work
# but break other important things
# Define these at you own risk
#
# $GLOBAL:TERM = "msys"
# $ENV:TERM = $GLOBAL:TERM
# $GLOBAL:TERM = "msys"
# $ENV:TERM = $GLOBAL:TERM
# $GLOBAL:TERMCAP = "C:/PROGRA~3/Git/etc/termcap"
# $ENV:TERMCAP = $GLOBAL:TERMCAP
# $GLOBAL:TERMINFO = "C:/PROGRA~3/Git/etc/terminfo"
# $ENV:TERMINFO = $GLOBAL:TERMINFO

$GLOBAL:vvvIsVerbose = $true
$GLOBAL:vimrc_base_name = "vimrc4me.vim"
$GLOBAL:vimrc_base_old = "pk_vimrc.vim"
$GLOBAL:netLcl = bs2s "$netLclDsk/00"


$GLOBAL:ErrorActionPreference = "SilentlyContinue"
$GLOBAL:DebugPreference = "SilentlyContinue"
try {
    $GLOBAL:homeNet = "C:/drpbx/Dropbox/$(hostname)"
    if ( Test-path -path "C:/drpbx/Dropbox" -PathType container ) {
      if ( -not ( Test-path -path $homeNet -PathType container ) ) {
          new-item -type directory -path $homeNet
      }
      if ( -not ( Test-path -path "$homeNet/temp" -PathType container ) ) {
          new-item -type directory -path "$homeNet/temp"
      }
    }
    else {
        throw
    }
}
catch {
    write-host -foregroundcolor magenta "**ERROR: [ cannot create $homeNet ] :RORRE**"
}

try {
    $GLOBAL:psDiaryNetD = bs2s "$homeNet/diary"
    if ( Test-path -path $GLOBAL:homeNet -PathType container ) {
      if ( -not ( Test-path -path $psDiaryNetD -PathType container ) ) {
          new-item -type directory -path $psDiaryNetD
      }
      if ( -not ( Test-path -path "$psDiaryNetD/temp" -PathType container ) ) {
          new-item -type directory -path "$psDiaryNetD/temp"
      }
    }
    else {
        throw
    }
}
catch {
    write-host -foregroundcolor magenta "**ERROR: [ cannot create $homeNet/diary ] :RORRE**"
}

try {
    $GLOBAL:sundryNetD = bs2s "$homeNet/sundry"
    if ( Test-path -path $GLOBAL:homeNet -PathType container ) {
      if ( -not ( Test-path -path $sundryNetD -PathType container ) ) {
          new-item -type directory -path $sundryNetD
      }
      if ( -not ( Test-path -path "$sundryNetD/temp" -PathType container ) ) {
          new-item -type directory -path "$sundryNetD/temp"
      }
    }
    else {
        throw
    }
}
catch {
    write-host -foregroundcolor magenta "**ERROR: [ cannot create $homeNet/sundry ] :RORRE**"
}
$GLOBAL:ErrorActionPreference = "Inquire"
$GLOBAL:DebugPreference = "Inquire"

$GLOBAL:psprofilBkNetD = bs2s "$homeNet/keep"
if ( Test-path -path $GLOBAL:homeNet -PathType container ) {
  if ( -not ( Test-path -path $GLOBAL:psprofilBkNetD -PathType container ) ) {
  new-item -type directory -path $GLOBAL:psprofilBkNetD
  }
}

$GLOBAL:psprofilBkLclD = bs2s "$homeLcl/keep"
if ( -not ( Test-path -path $GLOBAL:psprofilBkLclD -PathType container ) ) {
  new-item -type directory -path $GLOBAL:psprofilBkLclD
}

$GLOBAL:vimProfileNetD = bs2s $homeNet
$GLOBAL:vimProfileLclD = bs2s $homeLcl

$GLOBAL:hDriveMountPt = bs2s "//$sambahost"

$GLOBAL:tulavnecot = bs2s "$hDriveMountPt/RCS/tulavnecot.7z"
$GLOBAL:GVtsprefix = usrHostTStampPid

switch -regex ( $localhost ) { # {{{
  '(?six: \A (FultonJSheen) | (Desales) | (Azariah) | (Magdalene) \z )' {
    $GLOBAL:transcriptdir = bs2s "$homeLcl/scrolls"
    if ( -not ( Test-Path -path $transcriptdir -PathType container ) ) {
      New-Item -itemtype directory -path $transcriptdir
    }
    $GLOBAL:sundryTmpD = bs2s "$sundryLclD/temp"
    if ( -not ( Test-Path -path $sundryTmpD -PathType container ) ) {
      New-Item -itemtype directory -path $sundryTmpD
    }
#    $GLOBAL:cmdHistoryTmpD = bs2s "$cmdHistoryLclD/temp"
    if ( -not ( $nullstr -eq $vimProfileLclD ) ) {
        if ( -not ( Test-Path -path $vimProfileLclD -PathType container ) ) {
          New-Item -itemtype directory -path $vimProfileLclD
        }
        $GLOBAL:vimProfileDir = bs2s $vimProfileLclD
    }
    else {
        echo "vimProfileLclD is null"
    }
    if ( -not ( Test-Path -path "$homeLcl/vimbk" -PathType container ) ) {
      New-Item -itemtype directory -path "$homeLcl/vimbk"
    }
    sow vimBackupDir "$homeLcl/vimbk"
    sow vimbk "$vimBackupDirL"
  }
  default {
    if ( -not ( Test-Path -path "$homeNet/scrolls" -PathType container ) ) {
      New-Item -itemtype directory -path "$homeNet/scrolls"
    }
    $GLOBAL:transcriptdir = bs2s "$homeNet/scrolls"
    $GLOBAL:sundryTmpD = $env:temp
    $GLOBAL:psDiaryRecent = $env:temp
    $GLOBAL:vimProfileDir = bs2s $homeNet
    sow vimBackupDir "$homeNet/vimbk"
  }
} # }}}
if ( -not ( Test-Path -path $ENV:vimBackupDir -PathType container ) ) {
  New-Item -itemtype directory -path $ENV:vimBackupDir
}
Set-ItemProperty -Path HKCU:\Environment -Name vimBackupDir -Value $( s2bs $ENV:vimBackupDir )

$GLOBAL:cdTrail = "$transcriptdir/$localhost.cdTrail.txt"
$GLOBAL:cdJournal = "$transcriptdir/$localhost.cdJournal.txt"

touch $cdJournal

$GLOBAL:expiredFiles = bs2s "$sundryTmpD/$( usrHost ).probably.ok2nuke.x.txt"
$s32 = $($env:comspec -replace '.cmd[.]exe$', $nullstr)

sow docset  "C:/Documents and Settings"
sow win     "$env:SystemRoot"
sow sys64   "$winL/SysWOW64"
sow sys32   "$s32"
sow dbx     "$homeLclDsk/drpbx/Dropbox"
sow loh     "$homeLclDsk/drpbx/Dropbox/LOH"
sow pfd     "C:/Program Files (x86)"
sow pfx     "C:/Program Files"
sow win     "C:/Windows"
sow spc     "C:/strwbry/c"
sow sby     "C:/strwbry/perl"
sow mgw     "C:/MinGW"
sow mgw     "C:/mingw64"
sow sqlite  "C:/sqlite"
sow droid   "C:/bin/Android"
sow bin     "C:/bin"
sow din     "C:/users/ksk/bin"
sow pyd     "C:/Py35"
sow j7d     "C:/jdk7/w64/bin"
sow ekd     "C:/eklypz/w64"
sow dyd     "C:/Android/android-sdk"
sow abw     "C:/adt-bundle-windows-x86_64-20130219"
sow xbn     "C:/system/xbin"
sow adb     'C:/adb'
sow oracle  'C:/oracle'
sow prgdat  'C:/ProgramData'
sow PYTHONHOME 'C:/Py35'
$ENV:dbx = $dbxL


switch -regex ( $localhost ) { # {{{
  '(?six: \A (FultonJSheen) | (Desales) | (Azariah) | (Magdalene) \z )' {
    #sow vimLcl ( ,"$homeLclDsk/nomad" )
    sow VIM "$homeLclDsk/bin/vim"
    sow VIMRUNTIME "$VIM/$vimVersion"
  }

  default {
    #sow VIMLCL ( ,"$homeLcl" )
    sow VIM "$pfd/Vim"
    sow VIMRUNTIME "$VIM/$vimVersion"
  }
} # }}}


$GLOBAL:vimprofile = bs2s "$vimProfileDir/$vimrc_base_name"
$GLOBAL:vimprofold = bs2s "$vimProfileDir/$vimrc_base_old"
$GLOBAL:gooeyvimexe = bs2s "$VIM/$vimVersion/gvim.exe"
$GLOBAL:gooeyvim = bs2s "$gooeyvimexe -u $vimprofile"
$GLOBAL:crunchyvimexe = bs2s "$VIM/$vimVersion/vim.exe"
$GLOBAL:crunchyvim = bs2s "$crunchyvimexe -u $vimprofile"
$GLOBAL:bzipexe = shortnamequiet "$git/bzip2.exe"

ePlant git_editor $gooeyvimexe

sow JAVA_HOME "$bin/jdk1.8.0_162"
sow JRE_HOME  "$JAVA_HOME/jre"
sow CLASSPATH $(sunder "$JAVA_HOME/lib;$JRE_HOME/lib")
sow JRE8_64   $JRE_HOME
sow SMARTGIT_JAVA_HOME "$JAVA_HOME"

initiateTranscript $transcriptdir
ssinit

Register-EngineEvent PowerShell.Exiting {
    # Powershell_Exit_Processing Make sure that
    # anything you put here runs quickly
} -SupportEvent



pkpath
pypath
ePlant PYTHONDEBUG 1
ePlant PYTHONIOENCODING 'utf-8'
ePlant pathext "$env:pathext;.PL;.PY"
$dstack = @{}
setdir $GLOBAL:mmm

if ( -not ( Test-Path -path $psDiaryXmlFile -PathType Leaf ) )
{ # {{{
    generateDiary
} # }}}
loadAllSavedCommands
Set-PSReadlineOption -TokenKind String -ForegroundColor Cyan
$GLOBAL:ErrorActionPreference = "SilentlyContinue"
$GLOBAL:DebugPreference = "SilentlyContinue"
Import-Module $bin/posh-gvm/posh-gvm.psm1
$GLOBAL:ErrorActionPreference = "Inquire"
$GLOBAL:DebugPreference = "Inquire"
$GLOBAL:ACTIVATE_PROMPT = $true