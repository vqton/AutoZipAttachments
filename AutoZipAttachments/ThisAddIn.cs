using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

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
        
        private void Application_ItemSend(object Item, ref bool cancel)
        {
            Outlook.MailItem mailItem = Item as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.Recipients.Count > 0)
                {
                    if (mailItem.Recipients.Count > 0)
                    {
                        _emailSender.AddCC(mailItem);

                        // Add the backup group to the BCC field
                        _emailSender.AddBCC(mailItem, "Backup Group");
                    }
                }
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
