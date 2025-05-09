﻿<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#"%>
<SharePoint:RenderingTemplate ID="FileUploadControlTemplate" runat="server">
<Template>
<asp:FileUpload ID="UploadFileControl" runat="server" CssClass="ms-ButtonHeightWidth" Width="250px" />
<asp:Button ID="UploadButton" runat="server" CssClass="ms-ButtonHeightWidth" Width="50px" CausesValidation="false" Text="Upload"/>
<asp:Label ID="StatusLabel" runat="server" Width="300px" />
<asp:HiddenField ID="hdnFileName" runat="server" />
</Template>
</SharePoint:RenderingTemplate>
