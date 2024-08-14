@ECHO ON
MKDIR C:\FireEyeInstall
COPY "\\uson.local\netlogon\FireEyeHX\IMAGE_HX_AGENT_WIN_29.7.9\xagtSetup_29.7.9_universal.msi" C:\FireEyeInstall
COPY "\\uson.local\netlogon\FireEyeHX\IMAGE_HX_AGENT_WIN_29.7.9\agent_config.json" C:\FireEyeInstall
CD C:\FireEyeInstall
START /WAIT msiexec /i xagtSetup_29.7.9_universal.msi /qn CONFJSON=C:\FireEyeInstall\agent_config.json
