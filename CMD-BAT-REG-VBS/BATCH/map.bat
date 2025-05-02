net use k: "<insert file share>"
k:
if exist k:\"%username%" goto end1
mkdir "%username%"

:end1
c:
net use /delete k: /Y
net use k: /persistent:no "<insert file share>"