##########################################################################
# Joe Martin
# Cisco Systems, Inc.
# Basic UCS Configuration Data File for UCS Base Configuration Builder
# 5/11/14
# System: <Provide system description info here>
#
# Code provided as-is.  No warranty implied or included.
# This code is for example use only and not for production
#
# This script builds a new UCS system up with a standard configuration
# based on naming and features used by me.
#
# Below you will see a series of values to provide for the script to
# provide customized capabilities for your deployment
#
# Use the examples to get an idea of data to provide
#
##########################################################################
# IF NOT USING A SECTION, FILL THE ENTRIES WITH "".  Example: $FCPort = ""
##########################################################################
# BE CAREFUL NOT TO MAKE FIELD NAMES TOO LONG OR TO USE INVALID CHARACTERS
# OR YOUR BUILD MAY FAIL.
##########################################################################
Write-Host -ForegroundColor White -BackgroundColor DarkMagenta "Data File Module Loaded"

## Identify whether the system being accessed is an emulator or real system.
#If running against emulator "y", if running against a real UCS "n"
#By turning this on, the system will not re-ack the chassis' or FEX's which will cause the emulator to fail.  It will also not configure any server ports as this also causes the emulator to fail
$UCSEmulator = "y"

## UCS System IP or Host Name (VIP address for UCS Management)
# Example: $myucs = "9.9.9.9"
$myucs = "x.x.x.x"

## UCS Domain Number (Unique 16 bit Hex number to identify this UCS system as unique among other systems in the organization)
#This element is used to create the UUID, MAC, WWPNs and WWNNs addresses
#Info: MAC Addresses - 00:25:B4:XX:C1:00 - FF = Where XX = UCS Domain
#Info: UUID Suffixs - 00XX- = Where XX = UCS Domain
#Info: WWPN/WWNNs - 20:00:00:25:B5:XX:AA:00 - FF = Where XX = UCS Domain
#Info: IQN Suffix - UCSXX = Where XX = UCS Domain
#While UCSDomain is not a real UCS parameter is it used to ensure that addresses are not overlapped in the environment
# Example: $UCSDomain = "01"
#Max: 2 digit number
$UCSDomain = "?"

## Boot Policy
#Set ("y" or "n") for each boot from option to determine how blades will boot.  Can select multiple for yes(y) or no(n).
# Example: $BootFromSAN = "y" or "n"
$BootFromHD = "?"
$BootFromSAN = "?"
$BootFromiSCSI = "?"

## Chassis Power Policy
# Example: $ChassisPower = "grid" or "n+1"
$ChassisPower = "grid"

## Server Ports (Add or remove ports as needed.  Make sure to update last line with new QTY)
# Example: $ServerPort1 = @{Port = "1";	Slot = "1"; LabelA = "Chassis1A"; LabelB = "Chassis1B" }
#The use of Labels has not proven to be completely reliable in UCS.  Use at your own risk
#Labels Max: 16 Characters
$ServerPort1 = @{Port = "?";	Slot = "?"; LabelA = ""; LabelB = "" }
$ServerPort2 = @{Port = "?";	Slot = "?"; LabelA = ""; LabelB = "" }
#Make sure to match the entries in the array to the hash table
$ServerPort = @($ServerPort1, $ServerPort2)

## Uplink Ports(Add or remove ports as needed.  Make sure to update last line with new QTY)
# Example: $UplinkPort1 = @{Port = "32";  Slot = "1"; LabelA = "Nexus1-1-6"; LabelB = Nexus2-1-6" }
#The use of Labels has not proven to be completely reliable in UCS.  Use at your own risk
#Labels Max: 16 Characters
$UplinkPort1 = @{Port = "?";	Slot = "?"; LabelA = ""; LabelB = "" }
$UplinkPort2 = @{Port = "?";	Slot = "?"; LabelA = ""; LabelB = "" }
#Make sure to match the entries in the array to the hash table
$UplinkPort = @($UplinkPort1, $UplinkPort2)

##LAN Port Channels
#If using LAN Port Channels, set the value to "y"
#Value: y or n
$LANPortChannels       = "?"
#Fabric A LAN Port Channel Name
#Max: 16 Characters
$LANPortChannelAName   = "?"
#Fabric A LAN Port Channel Number
#Range: 1 - 256
$LANPortChannelANumber = "?"
#Fabric B LAN Port Channel Name
#Max: 16 Characters
$LANPortChannelBName   = "?"
#Fabric B LAN Port Channel Number
#Range: 1 - 256
$LANPortChannelBNumber = "?"

## Fibre Channel Ports(Add or remove ports as needed.  Make sure to update last line with new QTY)
#Must be an even number.  Must start from highest ports (1/32 down or 2/16 down)
# Example: $FCPort1 = @{Port = "15";   Slot = "2"; LabelA = "MDS1-1-6"; LabelB = "MDS2-1-6" }
# Example: $FCPort2 = @{Port = "16";   Slot = "2"; LabelA = "MDS1-1-7"; LabelB = "MDS2-1-7" }
#The use of Labels has not proven to be completely reliable in UCS.  Use at your own risk
#Labels Max: 16 Characters
$FCPort1 = @{Port = "?";	Slot = "?"; LabelA = ""; LabelB = "" }
$FCPort2 = @{Port = "?";	Slot = "?"; LabelA = ""; LabelB = "" }
#Make sure to match the entries in the array to the hash table
$FCPort = @($FCPort1, $FCPort2)

##SAN Port Channels
#If using SAN Port Channels, set the value to "y"
$SANPortChannels       = "?"
#Fabric A SAN Port Channel Name. 16 characters max, no spaces or special characters except '-' and '_'
#Max: 16 Characters
$SANPortChannelAName   = "?"
#Fabric A SAN Port Channel Number
#Range: 1 - 256
$SANPortChannelANumber = "?"
#Fabric B SAN Port Channel Name. 16 characters max, no spaces or special characters except '-' and '_'
#Max: 16 Characters
$SANPortChannelBName   = "?"
#Fabric B SAN Port Channel Number
#Range: 1 - 256
$SANPortChannelBNumber = "?"

## VSAN Info
#vSAN for Fabric A
# Example: $VSANidA = "12"
#Range: 1 - 4093
$VSANidA   = "?"
#FCOE vLAN for Fabric A
# Example: $FcoeVlanA = "1012"
#Range: 1 - 4093
$FcoeVlanA = "?"
#vSAN for Fabric B
# Example: $VSANidB = "11"
#Range: 1 - 4093
$VSANidB   = "?"
#FCOE vLAN for Fabric B
# Example: $FcoeVlanB = "1011"
#Range: 1 - 4093
$FcoeVlanB = "?"

## Name of FC HBAs (I prefer vHBA_A and vHBA_B)
# Example: $VHBAnameA = "vHBA_A"
#Max: 16 Characters
$VHBAnameA = "?"
$VHBAnameB = "?"

## SAN Controller Ports available for booting
#Use if supporting Boot from FC SAN
#Controller = A or B (On Controller)
#Port = #
#Count = 1 - X
#Fabric = A or B (On Fabric)
#Name = Controller Name
#WWPN = Target WWPN
# Example: $ArrayPort1 = @{Controller = "A";	Port = "1";	Count = "1";	Fabric = "A";	Name = "A-FP1";	WWPN = "50:00:AB:CD:EF:01:02:98" }
$ArrayPort1 = @{Controller = "?";	Port = "?";	Count = "1";	Fabric = "A";	Name = "?";	WWPN = "?" }
$ArrayPort2 = @{Controller = "?";	Port = "?"; 	Count = "2";	Fabric = "B";	Name = "?";	WWPN = "?" }
$ArrayPort3 = @{Controller = "?";	Port = "?"; 	Count = "3";	Fabric = "A";	Name = "?";	WWPN = "?" }
$ArrayPort4 = @{Controller = "?";	Port = "?"; 	Count = "4";	Fabric = "B";	Name = "?";	WWPN = "?" }
$ArrayPort5 = @{Controller = "?";	Port = "?"; 	Count = "5";	Fabric = "A";	Name = "?";	WWPN = "?" }
$ArrayPort6 = @{Controller = "?";	Port = "?"; 	Count = "6";	Fabric = "B";	Name = "?";	WWPN = "?" }
$ArrayPort7 = @{Controller = "?";	Port = "?"; 	Count = "7";	Fabric = "A";	Name = "?";	WWPN = "?" }
$ArrayPort8 = @{Controller = "?";	Port = "?"; 	Count = "8";	Fabric = "B";	Name = "?";	WWPN = "?" }
#Make sure to match the entries in the array to the hash table
$ArrayPort = @($ArrayPort1, $ArrayPort2, $ArrayPort3, $ArrayPort4, $ArrayPort5, $ArrayPort6, $ArrayPort7, $ArrayPort8)

## SAN Boot Matrix
#Name = Boot Policy Name and Service Profile Template Name. Max: 22 Characters
#APrimary, ASecondary, BPrimary, BSecondary = Count from $ArrayPort
#Use the 'Count' Field from the array port to create the matrix
#The names are also used for the Service Profile Templates
# Example: $BootMatrix1 = @{Name = "FC1526";	APrimary = "1";	ASecondary = "5";	BPrimary = "2";	BSecondary = "6" }
$BootMatrix1 = @{Name = "?";	APrimary = "1";	ASecondary = "5";	BPrimary = "2";	BSecondary = "6" }
$BootMatrix2 = @{Name = "?";	APrimary = "5";	ASecondary = "1";	BPrimary = "6";	BSecondary = "2" }
$BootMatrix3 = @{Name = "?";	APrimary = "3";	ASecondary = "7";	BPrimary = "4";	BSecondary = "8" }
$BootMatrix4 = @{Name = "?";	APrimary = "7";	ASecondary = "3";	BPrimary = "8";	BSecondary = "4" }
$BootMatrix5 = @{Name = "?";	APrimary = "2";	ASecondary = "6";	BPrimary = "1";	BSecondary = "5" }
$BootMatrix6 = @{Name = "?";	APrimary = "6";	ASecondary = "2";	BPrimary = "5";	BSecondary = "1" }
$BootMatrix7 = @{Name = "?";	APrimary = "4";	ASecondary = "8";	BPrimary = "3";	BSecondary = "7" }
$BootMatrix8 = @{Name = "?";	APrimary = "8";	ASecondary = "4";	BPrimary = "7";	BSecondary = "3" }
#Make sure to match the entries in the array to the hash table
$BootMatrix = @($BootMatrix1, $BootMatrix2, $BootMatrix3, $BootMatrix4, $BootMatrix5, $BootMatrix6, $BootMatrix7, $BootMatrix8)

## iSCSI vHBA Information
#iSCSI NIC's (I prefer iSCSI_A and iSCSI_B)
#Make sure to have created NIC's with these names in the VLAN/vNIC matrix below
# Example: $iSCSINicNameA = "iSCSI_A"
#Max: 16 Characters
$iSCSINicNameA = "?"
$iSCSINicNameB = "?"

## Use iSCSI IP Pool or DHCP for iSCSI Boot
#If using "Pool" then fill out the iSCSI IP Pools info below
# Example: $iSCSIPool = "Pool" or "DHCP"
$iSCSIPool = "?"


## Values entered into default iscsi-initiator-pool which cannot be deleted and I do not recommend using that pool for anything
$DefaultiSCSIPoolDefGW  = "1.1.1.1"
$DefaultiSCSIPrimDNS    = "1.1.1.2"
$DefaultiSCSISecDNS     = "1.1.1.3"
$DefaultiSCSIPoolFrom   = "1.1.1.4"
$DefaultiSCSIPoolTo     = "1.1.1.4"
$DefaultiSCSIPoolSubnet = "255.255.255.0"

## iSCSI IP Pool for Fabric A (For UCSM Assigned IP's for the iSCSI NIC.)
# Example: $DefaultiSCSIInitiatorPoolA = "iscsi-initiator-pool-A"
# Example: $iSCSIIPstartA              = "2.2.2.2"
# Example: $iSCSIIPendA                = "2.2.2.253"
# Example: $iSCSIDefGwA                = "2.2.2.1"
# Example: $iSCSISubnetA               = "255.255.255.0"
#If using DHCP then set the pool info to ""
#Max: 32 Characters - $DefaultiSCSIInitiatorPoolA
$DefaultiSCSIInitiatorPoolA = "?"
$iSCSIIPstartA              = "?"
$iSCSIIPendA                = "?"
$iSCSIDefGwA                = "?"
$iSCSISubnetA               = "?"

## iSCSI IP Pool for Fabric B (For UCSM Assigned IP's for the iSCSI NIC.)
# Example: $DefaultiSCSIInitiatorPoolB = "iscsi-initiator-pool-B"
# Example: $iSCSIIPstartB		    = "3.3.3.2"
# Example: $iSCSIIPendB 			    = "3.3.3.253"
# Example: $iSCSIDefGwB 			    = "3.3.3.1"
# Example: $iSCSISubnetB               = "255.255.255.0"
# If using DHCP then set the pool info to ""
#Max: 32 Characters - $DefaultiSCSIInitiatorPoolB
$DefaultiSCSIInitiatorPoolB = "?"
$iSCSIIPstartB              = "?"
$iSCSIIPendB                = "?"
$iSCSIDefGwB                = "?"
$iSCSISubnetB               = "?"

## iSCSI Target IP Address Information
# Example: $iSCSItargetA = "2.2.2.254"
# Example: $iSCSItargetB = "3.3.3.254"
# Example: $iSCSIiqn = "iqn.1991-05.com.microsoft:hd1-iscsi1-c-target"
#An iSCSI Target here is just for sample/example as you will need to provide unique targets for each service profile as you don't want to try to boot all your blades on a single IQN target LUN
#I recommend using an $iSCSIiqn = "CHANGEME" to remind you to change the IQN on each Service Profile
$iSCSItargetA = "?"
$iSCSItargetB = "?"
#Max: 256 Characters
$iSCSIiqn     = "?"

## IQN Pool
#Used for iSCSI Boot
# Example: $IQNPrefix = "iqn.1966-09.com.cisco.ucs"
#Max: 150 Characters
$IQNPrefix = "?"
# Example: $IQNSuffix = "UCS". I put the UCSDomain automatically on the end so in the quotes only put your text so the example ends up being UCS01 if the $UCSDomain = "01"
#Max: 62 Characters
$IQNSuffix = "?"
# Example: $IQNFrom = "0"
$IQNFrom   = "?"
# Example: $IQNTo = "255"
$IQNTo     = "?"

## Global QoS Settings and QoS Policies
#Select the global QoS settings to enable and their values
#Enabled - "y" or "n"
#CoS - 0 - 6
#Packet Drop - "drop" or "no-drop".  Only have one custom set to support packet drop
#Weight - none=(0), best-effort=(1), 2 - 10
#MTU - normal, fc, 1500 - 9216
#Multicast Optimized - "yes" or "no".  Only have one policy set to support multicast optimized
#
#Select the QoS Policy settings to enable and their values
#Burst - 1 - 65535
#Rate - line-rate or 1 - 9999999
#Host Control - none or full
#
######-Burst 10240 -HostControl "none" -Name "" -Prio "fc" -Rate "line-rate"
#Fibre Channel Global Defaults: Enabled=y(Always),	CoS=3,			Packet Drop=no-drop(Always),		Weight=5,		MTU=fc(Always),	Multicast Optimized=N/A(Always)
#Fibre Channel LAN Policy Defaults: 	Burst=10240, Rate="line-rate", HostControl = "none"
$FibreChannelQoSCoS 		   = "3"
$FibreChannelQoSWeight 		   = "5"
$FibreChannelQoSBurst		   = "10240"
$FibreChannelQoSRate		   = "line-rate"
$FibreChannelQoSHostControl	   = "none"
#Best Effort Global Defaults:   Enabled=yes(Always),	CoS=any(Always), 	Packet Drop=drop(Always),		Weight=5,		MTU=normal,     	Multicast Optimized=no
#BestEffort LAN Policy Defaults: 		Burst=10240, Rate="line-rate", HostControl = "none"
$BestEffortQoSWeight 		   = "5"
$BestEffortQoSMTU 			   = "normal"
$BestEffortQoSMulticastOptimized = "no"
$BestEffortQoSBurst		        = "10240"
$BestEffortQoSRate		        = "line-rate"
$BestEffortQoSHostControl	   = "none"
#Bronze Global Defaults: 		Enabled=no,		CoS=1,			Packet Drop=drop,				Weight=7,		MTU=normal, 	   	Multicast Optimized=no
#Bronze LAN Policy Defaults: 			Burst=10240, Rate="line-rate", HostControl = "none"
$BronzeQoSEnabled 		        = "n"
$BronzeQoSCoS 			        = "1"
$BronzeQoSPacketDrop 		   = "drop"
$BronzeQoSWeight 		        = "7"
$BronzeQoSMTU 			   	   = "normal"
$BronzeQoSMulticastOptimized     = "no"
$BronzeQoSBurst		        = "10240"
$BronzeQoSRate		        	   = "line-rate"
$BronzeQoSHostControl	        = "none"
#Silver Global Defaults: 		Enabled=no,		CoS=2,			Packet Drop=drop,				Weight=8,		MTU=normal, 		Multicast Optimized=no
#Silver LAN Policy Defaults: 			Burst=10240, Rate="line-rate", HostControl = "none"
$SilverQoSEnabled 		        = "n"
$SilverQoSCoS 			        = "2"
$SilverQoSPacketDrop 		   = "drop"
$SilverQoSWeight 		        = "8"
$SilverQoSMTU 			   	   = "normal"
$SilverQoSMulticastOptimized     = "no"
$SilverQoSBurst		        = "10240"
$SilverQoSRate		        	   = "line-rate"
$SilverQoSHostControl	        = "none"
#Gold Global Defaults: 			Enabled=no,			CoS=4,			Packet Drop=drop,			Weight=9,		MTU=normal, 		Multicast Optimized=no
#Gold LAN Policy Defaults: 			Burst=10240, Rate="line-rate", HostControl = "none"
$GoldQoSEnabled 		        = "n"
$GoldQoSCoS 			        = "4"
$GoldQoSPacketDrop 		        = "drop"
$GoldQoSWeight 		        = "9"
$GoldQoSMTU 			   	   = "normal"
$GoldQoSMulticastOptimized       = "no"
$GoldQoSBurst		       	   = "10240"
$GoldQoSRate		        	   = "line-rate"
$GoldQoSHostControl	      	   = "none"
#Platinum Global Defaults: 		Enabled=no,			CoS=5,			Packet Drop=no-drop,		Weight=10,	MTU=normal, 		Multicast Optimized=no
#Platinum LAN Policy Defaults: 		Burst=10240, Rate="line-rate", HostControl = "none"
$PlatinumQoSEnabled 		   = "n"
$PlatinumQoSCoS 			   = "5"
$PlatinumQoSPacketDrop 		   = "no-drop"
$PlatinumQoSWeight 		        = "10"
$PlatinumQoSMTU 			   = "normal"
$PlatinumQoSMulticastOptimized   = "no"
$PlatinumQoSBurst		        = "10240"
$PlatinumQoSRate		        = "line-rate"
$PlatinumQoSHostControl	        = "none"

## VLAN/vNIC Information (Follow the format below and add or remove entries as needed)
#vlanname = Human readable name for vlan
#vlannumber = numeric value for vlan
#macid = 2 digit hex number to be part of the mac address 00:25:B5:XX:YY:ZZ where XX = domainID and YY = macid and ZZ = Pool Range (00-ff)
#mtu = 1500 or 9000
#fabric = A-B, B-A, A or B. A-B means prefer fabric A but allow failover to fabric B. B-A means prefer fabric B but allow failover to fabric A. A means only use fabric A.  B means only use fabric B.
#QoSPolicy = "", "BestEffort", "Bronze", "Silver", "Gold", "Platinum"
# Example: $network1 = @{vlanname = "MGMT";				vlannumber = "22";	macid = "36"; 	mtu = "1500"; 	fabric = "A-B"; QoSPolicy = "BestEffort" }
#Max: 16 Characters - vlanname
#Regular vNICs with single VLAN
$network1  = @{vlanname = "?";	vlannumber = "?";		macid = "?"; 	mtu = "?"; 	fabric = "?"; QoSPolicy = "?" }
$network2  = @{vlanname = "?";	vlannumber = "?";		macid = "?"; 	mtu = "?"; 	fabric = "?"; QoSPolicy = "?" }
#Trunked vNICs with multiple VLANs
# Example: $network3  = @{ vlanname = "VM_Network";		vlannumber = "";		macid = "FF";	mtu = "9000";	fabric = "A-B"; QoSPolicy = "BestEffort" }
# Example: $network4  = @{ vlanname = "trunk10";			vlannumber = "10";		macid = "";	mtu = "9000";		fabric = "NONE";		trunknic = "VM_Network" }
# Example: $network5  = @{ vlanname = "trunk11";			vlannumber = "11";		macid = "";	mtu = "9000";		fabric = "NONE";		trunknic = "VM_Network"; 	defaultvlan = "y" }
$network3  = @{ vlanname = "?";	vlannumber = "?";		macid = "?";	mtu = "?";		fabric = "?";		QoSPolicy = "?" }
$network4  = @{ vlanname = "?";	vlannumber = "?";		macid = "?";	mtu = "?";		fabric = "?";		trunknic = "?" }
$network5  = @{ vlanname = "?";	vlannumber = "?";		macid = "?";	mtu = "?";		fabric = "?";		trunknic = "?"; 	defaultvlan = "?" }
#NICs that are non-failover sharing the same VLAN for deployments like VMWare where fabric failover is not recommended
# Example: $network15 = @{ vlanname = "The_VLAN";			vlannumber = "19";		macid = "";	mtu = "";		fabric = "NONE"; 									QoSPolicy = "" }
# Example: $network16 = @{ vlanname = "NIC_on_A";			vlannumber = "";		macid = "0A";	mtu = "9000";	fabric = "A";	trunknic = ""; 		defaultvlan = ""; 	QoSPolicy = "BestEffort" }
# Example: $network17 = @{ vlanname = "NIC_on_B";			vlannumber = "";		macid = "0B";	mtu = "9000";	fabric = "B";	trunknic = ""; 		defaultvlan = ""; 	QoSPolicy = "BestEffort" }
# Example: $network18 = @{ vlanname = "The_VLAN";			vlannumber = "";		macid = "";	mtu = "9000";	fabric = "A";	trunknic = "NIC_on_A"; 	defaultvlan = "y"; 	QoSPolicy = "BestEffort" }
# Example: $network19 = @{ vlanname = "The_VLAN";			vlannumber = "";		macid = "";	mtu = "9000";	fabric = "B";	trunknic = "NIC_on_B"; 	defaultvlan = "y"; 	QoSPolicy = "BestEffort" }
$network6 = @{ vlanname = "?";			vlannumber = "?";		macid = "?";	mtu = "?";	fabric = "?";	QoSPolicy = "?" }
$network7 = @{ vlanname = "?";			vlannumber = "?";		macid = "?";	mtu = "?";	fabric = "?";	trunknic = "?"; 	defaultvlan = "?"; QoSPolicy = "?" }
$network8 = @{ vlanname = "?";			vlannumber = "?";		macid = "?";	mtu = "?";	fabric = "?";	trunknic = "?"; 	defaultvlan = "?"; QoSPolicy = "?" }
$network9 = @{ vlanname = "?";			vlannumber = "?";		macid = "?";	mtu = "?";	fabric = "?";	trunknic = "?"; 	defaultvlan = "?"; QoSPolicy = "?" }
$network10= @{ vlanname = "?";			vlannumber = "?";		macid = "?";	mtu = "?";	fabric = "?";	trunknic = "?"; 	defaultvlan = "?"; QoSPolicy = "?" }
#Make sure to match the entries in the network array to the VLAN hash table
$network = @($network1, $network2, $network3, $network4, $network5, $network6, $network7, $network8, $network9, $network10)

## Native VLAN Info
#This is the native VLAN for northbound traffic...LAN Uplink traffic
# Example: $NativeVLANname = "default"
#default is the default native vlan name in the system
$NativeVLANname   = "?"
# Example: $NativeVLANnumber = "1"
#1 is the default native vlan number in the system
$NativeVLANnumber = "?"

## Enter the name of the NIC that the system will use for PXE boot in your boot policy.
# Example: $PXENIC = "PXE"
#If not using PXE boot then set value to ""
$PXENIC = "?"

## IPMI Access Policy (This password should be updated in UCSM after initial configuration is complete)
# Example: $IPMIpassword = "letmein"
#I highly recommend that your customer change this password for security and that this password be used just for the initial build and testing
$IPMIpassword = "?"

## Multicast Policy (This is the Multicast Querier IP Address that UCS will use)
# Example: $QuerierIpAddr = "9.9.9.1"
#Enter the Multicast querier address or set to "" if not using a multicast querier
$QuerierIpAddr = "?"

## Management Pool (This is systems management information)
#Default Gateway IP Address for UCSM Management
# Example: $DefGw = "9.9.9.1"
$DefGw = "?"

## Primary DNS IP Address for UCSM Management
# Example: $PriDNS = "9.9.9.10"
$PriDNS = "?"

## Secondary DNS IP Address for UCSM Management
# Example: $SecDNS = "9.9.9.11"
$SecDNS = "?"

## First IP Address in Blade Management Pool. (Must be on the same subnet as UCSM's management IP's)
# Example: $MgmtIPstart = "9.9.9.21"
$MgmtIPstart = "?"

## Last IP Address in Blade Management Pool. (Must be on the same subnet as UCSM's management IP's)
# Example: $MgmtIPend = "9.9.9.254"
$MgmtIPend = "?"

## Management Info
# Example: $UCSDesc = "UCS system in my lab"
#Max: 256 Characters
$UCSDesc      = "?"
# Example: $UCSOwner = "John Smith"
#Max: 32 Characters
$UCSOwner     = "?"
# Example: $UCSSite = "Datacenter 12 - anytown - Country"
#Max: 32 Characters
$UCSSite      = "?"
# Example: $UCSDNSDomain = "domain.com"
#Max: 256 Characters
$UCSDNSDomain = "?"
# Example: $SystemName = "MyUCS".  If left with just "" then the script will not change the default name assigned at startup
#Max: 30 Characters
$SystemName   = ""

## Timezone Management
#Must be full description such as "America/Los_Angeles (Pacific Time)".  You may have to log into a UCS system to find out the proper format for your timezone.
# Example: $Timezone = "America/Los_Angeles (Pacific Time)"
$Timezone = "?"

##NTP Servers
# Example: $NTPName = @("4.4.4.4", "5.5.5.5")
#Max: 64 Characters. IP Address or host.domain
$NTPName = @("?", "?")

## Important default variables and values
#Default Scub Policy
#Options: No_Scrub, BIOS_Scrub, Disk_Scrub, Full_Scrub
$DefaultScrub = "No_Scrub"
#Default User Acknowledgement of new Chassis and Rack Servers
#Options: user-acknowledged, immediate
$DefaultRackServerDiscovery = "immediate"
#Default Rack Management Connectcion Policy
#Options: user-acknowledged, auto-acknowledged
$DefaultRackManagement = "auto-acknowledged"
#Default Chassis/FEX Discovery - Action
#Options: 1-link, 2-link, 4-link, 8-link, platform-max
$DefaultDiscoveryAction = "1-link"
#Default Chassis/FEX Discovery - Link Grouping Preference
#Options: none, port-channel
$DefaultLinkGrouping = "port-channel"
#Default Power Control
#Options: default, No_Cap
$DefaultPowerControl = "No_Cap"
#Default Serial over LAN
#Options: No_SoL, SoL_9600, SoL_19200, SoL_38400, SoL_57600, SoL_115200
$DefaultSoL = "No_SoL"
#Default LAN Connectivity
#Options: Create your own logical name
$DefaultLANConnectivity = "vNICs"
#Options: Create your own logical name
$DefaultLANwiSCSIConnectivity = "vNICs_iSCSI"
#Default LAN Adapter Policy
#Options: "", Linux, SRIOV, Solaris, VMWare, VMWarePassThru, Windows, default
$DefaultLANAdapter = "Windows"
#Default vSAN Names
#Options: Create your own logical name and _A or _B will be appended to the end. ie vSAN_A and vSAN_B
$DefaultvSANName = "vSAN"
#Default HBA Connectivity
#Options: Create your own logical name
$DefaultHBAConnectivity = "vHBAs"
#Default SAN WWPN Pool Names
#Options: Create your own logical name and _A or _B will be appended to the end. ie Fabric_A and Fabric_B
$DefaultWWPNPool = "Fabric"

#Default Pool formats

#UUID Suffix
#Default: $UUIDfrom = "00"+$UCSDomain+"-000000000000"
#Default: $UUIDto   = "00"+$UCSDomain+"-0000000000FF"
$UUIDfrom = "00"+$UCSDomain+"-000000000000"
$UUIDto   = "00"+$UCSDomain+"-0000000000FF"

#MAC Pools
#Default: $MACfrom = "00:25:B5:"+$UCSDomain+":ID:00"
#Default: $MACto   = "00:25:B5:"+$UCSDomain+":ID:FF"
#ID will be replaced in the script with the value from $networkcount["macid"]
$MACfrom = "00:25:B5:"+$UCSDomain+":ID:00"
$MACto   = "00:25:B5:"+$UCSDomain+":ID:FF"

#WWNN Pool
#Default: $WWNNfrom = "20:00:00:25:B5:"+$UCSDomain+":00:00"
#Default: $WWNNto   = "20:00:00:25:B5:"+$UCSDomain+":00:FF"
$WWNNfrom = "20:00:00:25:B5:"+$UCSDomain+":00:00"
$WWNNto   = "20:00:00:25:B5:"+$UCSDomain+":00:FF"

#WWPN Pool - Fabric A
#Default: $WWPNaFrom = "20:00:00:25:B5:"+$UCSDomain+":AA:00"
#Default: $WWPNaTo   = "20:00:00:25:B5:"+$UCSDomain+":AA:FF"
$WWPNaFrom = "20:00:00:25:B5:"+$UCSDomain+":AA:00"
$WWPNaTo   = "20:00:00:25:B5:"+$UCSDomain+":AA:FF"

#WWPN Pool - Fabric A
#Default: $WWPNbFrom = "20:00:00:25:B5:"+$UCSDomain+":BB:00"
#Default: $WWPNbTo   = "20:00:00:25:B5:"+$UCSDomain+":BB:FF"
$WWPNbFrom = "20:00:00:25:B5:"+$UCSDomain+":BB:00"
$WWPNbTo   = "20:00:00:25:B5:"+$UCSDomain+":BB:FF"

#Additional Notes to send to console upon script completion
$SpecialNotes = 
"
?
"

#Launch Customization Script.  Leave at $null if not using a custom script.  Put the script name in Quotes if using one.
# Example: $CustomScript = "Custom UCS Settings - BLANK.ps1"
$CustomScript = $null

##################################################END OF DATA FILE ##########################################################