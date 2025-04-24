using System;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;

namespace CustomSiteField
{
    public partial class FileUploadFieldEditControl : UserControl, IFieldEditor
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }


        FileUploadCustomField _field = null;

        public bool DisplayAsNewSection
        {
            get { return true; }
        }

        public void InitializeWithField(SPField field)
        {
            this._field = field as FileUploadCustomField;
        }

        public void OnSaveChange(SPField field, bool isNewField)
        {
            FileUploadCustomField myField = field as FileUploadCustomField;
            myField.UploadDocumentLibrary = ddlDocLibs.SelectedItem.Text;
        }

        protected override void CreateChildControls()
        {
            base.CreateChildControls();
            SPListCollection objLists = SPContext.Current.Web.Lists;
            foreach (SPList objList in objLists)
            {
                if (objList is SPDocumentLibrary)
                    ddlDocLibs.Items.Add(new System.Web.UI.WebControls.ListItem(objList.Title, objList.Title));
            }
            if (!IsPostBack && _field != null)
            {
                if (!String.IsNullOrEmpty(_field.UploadDocumentLibrary))
                    FindControlRecursive<DropDownList>(this, "ddlDocLibs").Items.FindByText(_field.UploadDocumentLibrary).Selected = true;
            }
        }

        protected T FindControlRecursive<T>(Control rootControl, String id)
        where T : Control
        {
            T retVal = null;
            if (rootControl.HasControls())
            {
                foreach (Control c in rootControl.Controls)
                {
                    if (c.ID == id) return (T)c;
                    retVal = FindControlRecursive<T>(c, id);
                    if (retVal != null) break;
                }
            }
            return retVal;
        }
    }
}

