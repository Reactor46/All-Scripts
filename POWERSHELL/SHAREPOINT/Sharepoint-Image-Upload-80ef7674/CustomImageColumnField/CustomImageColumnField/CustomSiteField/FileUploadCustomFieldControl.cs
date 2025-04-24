using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System.Web.UI.WebControls;
using Microsoft.SharePoint.Client;
using System.IO;
using System.Net;
using System.Web;


namespace CustomSiteField
{
    class FileUploadCustomFieldControl : BaseFieldControl
    {
        protected FileUpload UploadFileControl;
        protected Button UploadButton;
        protected Label StatusLabel;
        protected HiddenField hdnFileName;

        public override void Focus()
        {
            if (Field == null || this.ControlMode == SPControlMode.Display)
            { return; }
            EnsureChildControls();
            UploadFileControl.Focus();
        }
        public override object Value
        {
            get
            {
                EnsureChildControls();
                if (hdnFileName.Value != string.Empty)
                    return hdnFileName.Value;
                else if (UploadFileControl.PostedFile != null)
                {
                    string strFileName = UploadFileControl.PostedFile.FileName.Substring(UploadFileControl.PostedFile.FileName.LastIndexOf("\\") + 1);
                    return strFileName;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                EnsureChildControls();
                hdnFileName.Value = (string)this.ItemFieldValue;
                StatusLabel.Text = "File: <a href='" + (string)this.ItemFieldValue + "' target='_blank'>View (" + (string)this.ItemFieldValue + ")</a>";
            }
        }

        protected override void CreateChildControls()
        {
            if (Field == null || this.ControlMode == SPControlMode.Display)
            { return; }

            base.CreateChildControls();

            UploadFileControl = (FileUpload)TemplateContainer.FindControl("UploadFileControl");
            UploadButton = (Button)TemplateContainer.FindControl("UploadButton");
            StatusLabel = (Label)TemplateContainer.FindControl("StatusLabel");
            hdnFileName = (HiddenField)TemplateContainer.FindControl("hdnFileName");

            UploadButton.Click += new EventHandler(UploadButton_Click);

            if (hdnFileName.Value == string.Empty)
                StatusLabel.Text = "Select file and click on upload.";
            else
                StatusLabel.Text = "File: <a href='" + hdnFileName.ToString() + "' target='_blank'>View (" + hdnFileName.ToString() + ")</a>";
            Controls.Add(StatusLabel);
        }

        protected void UploadButton_Click(object sender, EventArgs e)
        {
            try
            {
                SPWeb sourceWeb;

                SPSite sourceSite = SPControl.GetContextSite(Context);
                sourceWeb = sourceSite.AllWebs["/"];
                sourceWeb.AllowUnsafeUpdates = true;
                FileUploadCustomField _field = (FileUploadCustomField)this.Field;
                SPList objList = sourceWeb.Lists[_field.UploadDocumentLibrary];
                SPFolder destFolder = objList.RootFolder;

                if (UploadFileControl.PostedFile == null) return;
                string strFileName = UploadFileControl.PostedFile.FileName.Substring(UploadFileControl.PostedFile.FileName.LastIndexOf("\\") + 1);

                Stream fStream = UploadFileControl.PostedFile.InputStream;

                SPFile objFile = destFolder.Files.Add(strFileName, fStream, true);
                objFile.Item.UpdateOverwriteVersion();

                StatusLabel.Text = "Upload File :: Success <a href='" + "/" + objFile.ParentFolder + "/" + strFileName + "' target='_blank'>View (" + "/" + objFile.ParentFolder + "/" + strFileName + ")</a>";
                hdnFileName.Value = "/" + objFile.ParentFolder + "/" + strFileName;
                sourceWeb.Dispose();
            }
            catch (Exception ex)
            {
                StatusLabel.Text = "Upload File :: Failed " + ex.Message;
            }

        }

        protected override string DefaultTemplateName
        {
            get
            {
                return "FileUploadControlTemplate";
            }
        }
    }
}
