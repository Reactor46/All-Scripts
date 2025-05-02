# https://hub.docker.com/r/microsoft/powershell/
# https://github.com/PowerShell/PowerShell-Docker
# https://github.com/PowerShell/PowerShell

FROM mcr.microsoft.com/powershell:latest
ADD VERSION .

# adds wget
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
	wget unzip

#RUN TEMP_DEB="$(mktemp)" \
#    && wget -O "$TEMP_DEB" 'https://github.com/PowerShell/PowerShell/releases/download/v6.1.1/powershell_6.1.1-1.ubuntu.18.04_amd64.deb' \
#    && dpkg -i "$TEMP_DEB" \
#    && rm -f "$TEMP_DEB"

# Copy UCS PowerTool Suite binaries from local system
RUN mkdir -p ~/.local/share/powershell/Modules
RUN mkdir -p ~/.config/powershell/

# Install Cisco UCS powertool
# ADD https://communities.cisco.com/servlet/JiveServlet/download/74217-2-149644/ucspowertoolcore.zip /tmp
ADD https://community.cisco.com/kxiwq67737/attachments/kxiwq67737/4354j-docs-cisco-dev-ucs-integ/518/2/ucspowertoolcore.zip /tmp
RUN unzip /tmp/ucspowertoolcore.zip -d ~/.local/share/powershell/Modules/
RUN mv ~/.local/share/powershell/Modules/Start-UcsPowerTool.ps1 ~/.config/powershell/Microsoft.PowerShell_profile.ps1 -f

# Use PowerShell as the default shell
# Use array to avoid Docker prepending /bin/sh -c
CMD [ "pwsh" ]
