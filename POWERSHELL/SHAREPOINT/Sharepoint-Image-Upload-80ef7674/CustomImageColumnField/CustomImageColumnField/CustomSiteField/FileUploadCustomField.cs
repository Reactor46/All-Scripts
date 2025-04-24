using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;

namespace CustomSiteField
{
    class FileUploadCustomField : SPFieldText
    {
        public FileUploadCustomField(SPFieldCollection fields, string fieldName) : base(fields, fieldName) { Init(); }
        public FileUploadCustomField(SPFieldCollection fields, string typeName, string displayName) : base(fields, typeName, displayName) { Init(); }

        private string _UploadDocumentLibrary = string.Empty;
        public string UploadDocumentLibrary
        {
            get
            {
                return _UploadDocumentLibrary;
            }
            set
            {
                this.SetCustomProperty("UploadDocumentLibrary", value);
                _UploadDocumentLibrary = value;
            }
        }

        private void Init()
        {
            this.UploadDocumentLibrary = this.GetCustomProperty("UploadDocumentLibrary") + string.Empty;
        }

        public override void Update()
        {
            this.SetCustomProperty("UploadDocumentLibrary", this.UploadDocumentLibrary);
            base.Update();
        }

        public override BaseFieldControl FieldRenderingControl
        {
            get
            {
                BaseFieldControl fieldControl = new FileUploadCustomFieldControl();
                fieldControl.FieldName = InternalName;
                return fieldControl;
            }
        }
    }
}
