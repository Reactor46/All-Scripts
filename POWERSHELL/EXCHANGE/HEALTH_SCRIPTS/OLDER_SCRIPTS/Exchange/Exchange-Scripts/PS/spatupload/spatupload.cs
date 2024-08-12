using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Diagnostics;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport.Email;
using Microsoft.Exchange.Data.Transport.Smtp;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Exchange.Data.Common;
using SharepointAttachmentUploadAgent.UploadwebService;

namespace msgdevExchangeRoutingAgents
{
    public class SharepointAttachmentUploadFactory : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            RoutingAgent emSpUpload = new SharepointAttachUploadAgent();
            return emSpUpload;
        }
    }
}
public class SharepointAttachUploadAgent : RoutingAgent
{
    public SharepointAttachUploadAgent()
    {
        base.OnSubmittedMessage += new SubmittedMessageEventHandler(SharepointAttachUploadAgent_OnSubmittedMessage);
    }

    void SharepointAttachUploadAgent_OnSubmittedMessage(SubmittedMessageEventSource source, QueuedMessageEventArgs e)
    {
        ArrayList dmarray = new ArrayList();
        dmarray.Add("domain.com");
        Boolean pmProcMessage = false;
        EmailMessage emMessage = e.MailItem.Message;
        foreach (EnvelopeRecipient recp in e.MailItem.Recipients) {
            if (dmarray.Contains(recp.Address.DomainPart)== false) {
                pmProcMessage = true;
            }
        
        }
        if (pmProcMessage == true) { ProcessMessage(emMessage); }
    }
    public static byte[] ReadFully(Stream stream, int initialLength)
    {
        // ref Function from http://www.yoda.arachsys.com/csharp/readbinary.html
        // If we've been passed an unhelpful initial length, just
        // use 32K.
        if (initialLength < 1)
        {
            initialLength = 32768;
        }

        byte[] buffer = new byte[initialLength];
        int read = 0;

        int chunk;
        while ((chunk = stream.Read(buffer, read, buffer.Length - read)) > 0)
        {
            read += chunk;

            // If we've reached the end of our buffer, check to see if there's
            // any more information
            if (read == buffer.Length)
            {
                int nextByte = stream.ReadByte();

                // End of stream? If so, we're done
                if (nextByte == -1)
                {
                    return buffer;
                }

                // Nope. Resize the buffer, put in the byte we've just
                // read, and continue
                byte[] newBuffer = new byte[buffer.Length * 2];
                Array.Copy(buffer, newBuffer, buffer.Length);
                newBuffer[read] = (byte)nextByte;
                buffer = newBuffer;
                read++;
            }
        }
        // Buffer is now too big. Shrink it.
        byte[] ret = new byte[read];
        Array.Copy(buffer, ret, read);
        return ret;
    }
    static void ProcessMessage(EmailMessage emEmailMessage)
    {
        for (int index = emEmailMessage.Attachments.Count - 1; index >= 0; index--)
        {
            Attachment atAttach = emEmailMessage.Attachments[index];
            if (atAttach.EmbeddedMessage == null)
            {
                if (atAttach.AttachmentType == AttachmentType.Regular & atAttach.FileName != null)
                {
                    // Find Any PDF attachments with Quote in the File Name
                    if (atAttach.FileName.Length >= 3)
                    {
                        String feFileExtension = atAttach.FileName.Substring((atAttach.FileName.Length - 4), 4);
                        if (feFileExtension.ToLower() == ".pdf" | atAttach.FileName.ToLower().IndexOf("quote") != -1)
                        {
                            Stream attachstream = atAttach.GetContentReadStream();
                            String uploadResult = uploadAttachment(attachstream, atAttach.FileName.ToString());
                        }

                    }
                }
                atAttach = null;
            }
        }
     }
    static string uploadAttachment(Stream atAttachStream,String fnFileName) {
        string nfNewFileName = "-Sent(" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")+").pdf";
        
        string urUploadResult;
        byte[] atBytes = ReadFully(atAttachStream, (int)atAttachStream.Length);
        Files spUploader = new Files();
        string spDocumentLibrary = "http://servername/sites/Quotes/Shared%20Documents";
        string strUserName = "username";
        string strPassword = "password";
        string strDomain = "domain";
        System.Net.CredentialCache spUploadUserCredentials = new System.Net.CredentialCache();
        spUploadUserCredentials.Add(new System.Uri(spDocumentLibrary),
           "NTLM",
           new System.Net.NetworkCredential(strUserName, strPassword, strDomain)
           );
        spUploader.PreAuthenticate = true;
        spUploader.Credentials = spUploadUserCredentials;
        urUploadResult = spUploader.UploadDocument(fnFileName.ToLower().Replace(".pdf", nfNewFileName), atBytes, spDocumentLibrary);
        return urUploadResult;

    
    }
}