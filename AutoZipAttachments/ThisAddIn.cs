using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace AutoZipAttachments
{
    public partial class ThisAddIn
    {
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
             Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);

            
        }

        private void NewInspectorHandler(Inspector inspector)
        {
            if (inspector.CurrentItem is Outlook.MailItem mailItem)
            {
                string body = mailItem.HTMLBody;
                // Set font properties for the email body
                mailItem.HTMLBody = $@"
                    <div style=""font-family: Arial; font-size: 14pt;"">
                         {body}
                    </div>";
                
            }
        }

       
        private void Application_ItemSend(object item, ref bool cancel)
        {

            if (item is Outlook.MailItem mailItem)
            {
                string outputPath = Path.GetTempPath(); // Define your output path here

                IEmailFormatter emailFormatter = new EmailFormatter();
                IEmailRecipientManager emailRecipientManager = new EmailRecipientManager();
                IAttachmentCompressor attachmentCompressor = new AttachmentCompressor();

                EmailSender emailSender = new EmailSender(emailFormatter, emailRecipientManager, attachmentCompressor);
                emailSender.CompressAttachments(mailItem, outputPath);
            }
        }




        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
