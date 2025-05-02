##############################################################################
# Home, Profile and Data Share Creation Script                               #
# For use on Windows Server 2008 R2 w/AD Users                               #
# Creates directories, shares directories, sets permissions with icacls      #
##############################################################################

# Make the directories
md e:\samba
md e:\samba\data
md e:\samba\dna
# md e:\samba\ais_data  ## This can be created later
# md e:\samba\planner   ## This can be created later
md e:\profiles
md e:\home
# Share the directories
net share Data=e:\samba\data /grant:everyone,FULL /CACHE:None
net share profiles$=e:\profiles /grant:everyone,FULL /CACHE:None ## $ hides the share from normal browsing
net share homes$=e:\home /grant:everyone,FULL /CACHE:None  ## $ hides the share from normal browsing
net share DNA=e:\samba\DNA /grant:everyone,FULL /CACHE:None
# net share AIS_DATA=e:\samba\ais_data /grant:everyone,FULL /CACHE:None  ## This can be shared later
# net share Planner=e:\samba\planner /grant:everyone,FULL /CACHE:None    ## This can be shared later

# Lets set the NTFS permissions on the directories for the Data Shares (Samba and Subfolders)
# “(OI)(CI):F” means Full Control “This Folder, Subfolders and files”
# “(OI)(CI):M” means Modify “This Folder, Subfolders and files”
# “/inheritance:r” means remove all inherited ACL’s from parent

icacls "e:\samba" /grant "domain admins":(OI)(CI)F /inheritance:r
icacls "e:\samba" /grant "everyone":(OI)(CI)M /inheritance:r

# Any directories set under the E:\Samba directory will have to have its permissions
# set appropriately for the group or OU or user that requires access.

# Lets set the NTFS permissions on the directories for the Profiles Share
# On the profiles share, only the “domain admins” should be allowed to enter all
# “Folders, Subfolders and files” (hence the (OI)(CI):F) , everyone else should be
# able to to ready “this folder only”. So without an combination of (CI) and/or (OI)
# it means “this folder only”

icacls "e:\profiles" /grant "domain admins":(OI)(CI)F /inheritance:r
icacls "e:\profiles" /grant "everyone":R /inheritance:r

# Upon creating a new user, the Domain Admin should manually create a profile folder
# for the user and add the user with appropriate rights. If this becomes corrupted
# it will need to be addressed. Examples: "User Signed in with Temporary Profile"

# Lets set the NTFS permissions on the directories for the User Home
# Same permissions as the User Profiles

icacls "e:\home" /grant "domain admins":(OI)(CI)F /inheritance:r
icacls "e:\home" /grant "everyone":R /inheritance:r
