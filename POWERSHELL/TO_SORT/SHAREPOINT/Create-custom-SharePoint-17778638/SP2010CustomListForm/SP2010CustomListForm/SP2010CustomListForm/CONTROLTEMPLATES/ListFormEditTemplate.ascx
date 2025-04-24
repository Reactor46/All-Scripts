<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls"
    Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages"
    Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ListFormEditTemplate.ascx.cs"
    Inherits="SP2010CustomListForm.CONTROLTEMPLATES.ListFormEditTemplate" %>
<%@ Register TagPrefix="wssuc" TagName="ToolBar" Src="~/_controltemplates/ToolBar.ascx" %>
<%@ Register TagPrefix="wssuc" TagName="ToolBarButton" Src="~/_controltemplates/ToolBarButton.ascx" %>
<%@ Register TagPrefix="myCustomForm" TagName="AddForm" Src="~/_controltemplates/ListFormEditRender.ascx" %>
<SharePoint:RenderingTemplate ID="ListFormEdit" runat="server">
    <Template>
        <span id='part1'>
            <SharePoint:InformationBar ID="InformationBar1" runat="server" />
            <div id="listFormToolBarTop">
                <wssuc:ToolBar CssClass="ms-formtoolbar" ID="toolBarTbltop" RightButtonSeparator="&amp;#160;"
                    runat="server">
                    <Template_RightButtons>
                        <SharePoint:NextPageButton ID="NextPageButton1" runat="server" />
                        <SharePoint:SaveButton ID="SaveButton1" runat="server" />
                        <SharePoint:GoBackButton ID="GoBackButton1" runat="server" />
                    </Template_RightButtons>
                </wssuc:ToolBar>
            </div>
            <SharePoint:FormToolBar ID="FormToolBar1" runat="server" />
            <SharePoint:ItemValidationFailedMessage ID="ItemValidationFailedMessage1" runat="server" />
            <table class="ms-formtable" style="margin-top: 8px;" border="0" cellpadding="0" cellspacing="0"
                width="100%">
                <SharePoint:ChangeContentType ID="ChangeContentType1" runat="server" />
                <SharePoint:FolderFormFields ID="FolderFormFields1" runat="server" />
                <!-- myCustomForm -->
                <myCustomForm:AddForm ID="AddForm1" runat="server" />
                <!-- myCustomForm -->
                <SharePoint:ListFieldIterator ID="ListFieldIterator1" runat="server" />
                <SharePoint:ApprovalStatus ID="ApprovalStatus1" runat="server" />
                <SharePoint:FormComponent ID="FormComponent1" TemplateName="AttachmentRows" runat="server" />
            </table>
            <table cellpadding="0" cellspacing="0" width="100%">
                <tr>
                    <td class="ms-formline">
                        <img src="/_layouts/images/blank.gif" width='1' height='1' alt="" />
                    </td>
                </tr>
            </table>
            <table cellpadding="0" cellspacing="0" width="100%" style="padding-top: 7px">
                <tr>
                    <td width="100%">
                        <SharePoint:ItemHiddenVersion ID="ItemHiddenVersion1" runat="server" />
                        <SharePoint:ParentInformationField ID="ParentInformationField1" runat="server" />
                        <SharePoint:InitContentType ID="InitContentType1" runat="server" />
                        <wssuc:ToolBar CssClass="ms-formtoolbar" ID="toolBarTbl" RightButtonSeparator="&amp;#160;"
                            runat="server">
                            <Template_Buttons>
                                <SharePoint:CreatedModifiedInfo ID="CreatedModifiedInfo1" runat="server" />
                            </Template_Buttons>
                            <Template_RightButtons>
                                <SharePoint:SaveButton ID="SaveButton2" runat="server" />
                                <SharePoint:GoBackButton ID="GoBackButton2" runat="server" />
                            </Template_RightButtons>
                        </wssuc:ToolBar>
                    </td>
                </tr>
            </table>
        </span>
        <SharePoint:AttachmentUpload ID="AttachmentUpload1" runat="server" />
    </Template>
</SharePoint:RenderingTemplate>
