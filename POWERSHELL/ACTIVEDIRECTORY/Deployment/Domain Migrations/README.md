# Deployment

The scripts in this folder are designed to automate deployment and migration of certain services. For the most part, this is all in the context of small to medium-sized business.

## Domain Migration

This script is designed to automate the migration of one domain controller to another. Please note that this is something that is still being tested. The goal is for it to run completely on it's own from a management workstation.

#### Useful Bits

- Windows Server 2012 R2 -- has not been tested in 2008R2 or 2012, but should work.
- You must set your DNS manually on everything but the domain controller with anything that has a static address.
- This script will not work if cannot contact your primary DC via Remote PowerShell from your soon-to-be domain controller. It will just stop.

#### Instructions

1. Edit the script. There are variables you will need to adjust at the top.
2. Run the script. You shouldn't be asked to do anything at any point.
3. After Domain Services is installed, you should adjust DNS on all devices that have DNS statically assigned. This includes your DHCP server, of course. At this point, your server will reboot.
4. After the reboot, re-run the script. 
