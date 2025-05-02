VMware Scanner v1.4
Created by: the.anykey@gmail.com
Date: 22 July 2011


What does the VMware Scanner do?
=================================
The VMware scanner is a highly multi threaded application to detect VMware servers (ESX, ESXi, VMware Server and VirtualCenter). The detection is done by scanning for open SSL ports (port 443) and then send the VMware API command to get the basic information of the server. Retrieving the product name, version number and build number can be done with no credentials.

Troubleshooting
===============
The scanner stores all found servers in a text file in the same directory of the scanner software. It also keeps a file called "lastip.txt" that tells you what the last IP that was scanned, so if the client crashes you know where to continue.

Not every windows OS can deal with hunderds of threads well (I have found), I would recommend to run the VMware scanner on a windows server operating system, as I found it can best deal with 750+ threads. If you are testing under a desktop windows OS, I would recommend not putting the max thread count to high (max 400) but this will slow down scanning rate.


Change History
==============

July 22 2011, v1.4 -Fix bug that ip range could not contain a zero.
July 19 2011, v1.3 -First public release


Comments
========
Any commments / questions, please email me at the.anykey@gmail.com or check out information on www.run-virtual.com


