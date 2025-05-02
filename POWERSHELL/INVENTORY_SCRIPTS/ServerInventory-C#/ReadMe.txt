How to use
----------

In case you support multiple domains
------------------------------------
Create an empty Database Inventory
Make sure Inventory.Server has a connection string to this database, and your database can be reached from you Inventory.Server

In case you have only one domain
--------------------------------
No need to create database
Comment out the line in InventoryContext() constructor


Firewall
--------
Add a Firewall rule to the Inventory.Server allowing incoming TCP on port 8080
Update InventoryServerURL setting of Inventory.Client (inventory\inventoryclient\app.config)

Run Server as Administrator
Run Client as Administrator on all machines you want to include in the Inventory (prerequisite: .net 4.5 Framework)



