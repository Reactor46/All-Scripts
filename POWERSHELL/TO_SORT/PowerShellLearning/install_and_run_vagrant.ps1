# Make sure you run PowerShell as an administrator and navigate to the directory where this script is stored
# Following that, execute the script with this command: .\install_and_run_vagrant.ps1

# Install Chocolatey (if not installed)
if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
    Set-ExecutionPolicy Bypass -Scope Process -Force
    iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
}

# Install VirtualBox and Vagrant using Chocolatey
choco install virtualbox -y
choco install vagrant -y

# Remove VBoxManage DHCP server for vboxnet0. 
# Vagrant sometimes throws up an error saying you are using such a server when you aren't upon spinning up.
& 'C:\Program Files\Oracle\VirtualBox\VBoxManage.exe' dhcpserver remove --netname HostInterfaceNetworking-vboxnet0

# Set the Vagrant project directory path (I just put it in the directory from which this script originates)
$vagrantProjectDir = "C:\Users\silad\Documents\Powershell_stuff\PowerShellLearning"

# Change to the Vagrant project directory
Set-Location $vagrantProjectDir

# Remove the existing Vagrantfile (if it exists)
if (Test-Path "Vagrantfile") {
    Remove-Item -Path "Vagrantfile" -Force
}

# Specify the Vagrant box you want to use (Ubuntu 20.04 in this example)
$vagrantBox = "generic/ubuntu2004"

# Run `vagrant init` only if Vagrantfile doesn't exist
if (-not (Test-Path "Vagrantfile")) {
    vagrant init $vagrantBox
}

# Start and provision the Vagrant virtual machine
vagrant up

# Connect to the Vagrant virtual machine via SSH. You will now be in an ubuntu environment.
vagrant ssh


