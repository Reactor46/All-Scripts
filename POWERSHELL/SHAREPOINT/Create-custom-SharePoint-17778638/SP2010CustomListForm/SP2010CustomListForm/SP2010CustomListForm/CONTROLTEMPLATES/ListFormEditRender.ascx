<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ListFormEditRender.ascx.cs" Inherits="SP2010CustomListForm.CONTROLTEMPLATES.ListFormEditRender" %>

<h1>
    Hi!!!!! My Custom List Template</h1>
<table width="100%" border="0" id="tblEditform" runat="server" cellpadding="0" cellspacing="0">
    <tr>
        <td valign="top">
            <table cellspacing="0" cellpadding="4" border="0" width="100%">
                <tr>
                    <td class="ms-vb">
                        &nbsp;
                    </td>
                </tr>
            </table>
            <table border="0" width="100%">
                <tr>
                    <td>
                        <table border="0" cellspacing="0" width="100%">
                            <tr>
                                <td class="ms-formlabel" valign="top" nowrap="true" width="35%">
                                    <b>Title:</b>
                                </td>
                                <td class="ms-formbody" valign="top">
                                    <SharePoint:FormField runat="server" ID="ff4" ControlMode="Edit" FieldName="Title" />
                                </td>
                            </tr>
                            <tr>
                                <td width="35%" class="ms-formlabel">
                                    <b>Custom:</b>
                                </td>
                                <td class="ms-formbody">
                                    <SharePoint:FormField runat="server" ID="MyCustom" ControlMode="Edit" FieldName="MyCustom" />
                                </td>
                            </tr>
                           
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
