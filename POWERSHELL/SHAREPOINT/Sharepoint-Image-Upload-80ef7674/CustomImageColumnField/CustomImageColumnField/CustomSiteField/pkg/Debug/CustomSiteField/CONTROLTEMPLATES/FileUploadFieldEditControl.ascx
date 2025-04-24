<%@ Assembly Name="CustomSiteField, Version=1.0.0.0, Culture=neutral, PublicKeyToken=918a4ed53d2b8c02" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" Inherits="CustomSiteField.FileUploadFieldEditControl" %>
<%@ Register TagPrefix="wssuc" TagName="InputFormControl" Src="~/_controltemplates/InputFormControl.ascx" %>
<%@ Register TagPrefix="wssuc" TagName="InputFormSection" Src="~/_controltemplates/InputFormSection.ascx" %>
<wssuc:InputFormSection runat="server" id="UploadControl" Title="Upload Control Configuration Section">
<template_inputformcontrols>
<wssuc:InputFormControl ID="InputFormControl1" runat="server" LabelText="Select document library">
<Template_Control>
<asp:DropDownList ID="ddlDocLibs" runat="server" CssClass="ms-ButtonheightWidth" Width="250px" />
</Template_Control>
</wssuc:InputFormControl>
</template_inputformcontrols>
</wssuc:InputFormSection>
