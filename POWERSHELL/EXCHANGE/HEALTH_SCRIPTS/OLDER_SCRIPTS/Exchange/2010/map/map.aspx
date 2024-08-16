<%@ Page Language="C#" AutoEventWireup="true"  CodeFile="map.aspx.cs" Inherits="_Map" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <script src="http://maps.google.com/maps?file=api&amp;v=2&amp;key=ABQIAAAAm-6QBXh6UHo7Yj0-PB__fhS9HErgP3HqyjjrP_9V7dnPNjbm1BROwdiPKB-Rj2hr4xTQe1udk1aOHA"
            type="text/javascript"></script>    
    <title>Untitled Page</title>
</head>
<body onload="initialize()">
    <form id="form1" runat="server">
    <div style="height: 123px">
    
        <asp:TextBox ID="AddressBox" runat="server" Height="122px" ReadOnly="True" 
            TextMode="MultiLine" Width="226px"></asp:TextBox>
    
    </div>
    </form><div id="map_canvas" style="width: 500px; height: 300px;"></div><div id="pano" style="width: 500px; height: 300px; left: 510px;"></div>
</body>
</html>
