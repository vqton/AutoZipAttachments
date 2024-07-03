using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace AutoZipAttachments
{
    public partial class ThisAddIn
    {
        private IEmailSender _emailSender;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Application.ItemSend += new ApplicationEvents_11_ItemSendEventHandler(CompressAttachments);
            _emailSender = new EmailSender();
            Application.Inspectors.NewInspector += NewInspectorHandler;

            Application.ItemSend += Application_ItemSend;
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

       
        private void Application_ItemSend(object Item, ref bool cancel)
        {
            Outlook.MailItem mailItem = Item as Outlook.MailItem;
            if (mailItem != null)
            {
                // Compress attachments before sending
                string outputPath = @"E:\temp"; // Define your desired output path
                _emailSender.CompressAttachments(mailItem, outputPath);
                if (mailItem.Recipients.Count > 0)
                {
                    
                    mailItem.Save();
                    if (mailItem.Recipients.Count > 0)
                    {
                        _emailSender.AddCC(mailItem);

                        // Add the backup group to the BCC field
                        string[] bccGroup = new string[] { "tonvqsgc@outlook.com", "tonqvu@gmail.com", "vuquangton@outlook.com", "vuquangton@ymail.com" };
                        _emailSender.AddBCC(mailItem, bccGroup);
                    }
                }
                //_emailSender.MoveTempDirectoryToSystemTemp(outputPath);
               _emailSender.FormatEmail(mailItem);
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
