using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.IO;
using System.IO.Compression;
using Microsoft.Office.Interop.Outlook;
namespace AutoZipAttachments
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.ItemSend += Application_ItemSend;
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            var mail = Item as MailItem;
            if (mail != null)
            {
                if (mail.Attachments.Count > 0)
                {
                    //var zipFile = Path.GetTempFileName() + ".zip";
                    var tempPath = Environment.GetEnvironmentVariable("temp");
                    var zipFile = Path.Combine(tempPath, "Archive.zip");
                    

                    using (var archive = ZipFile.Open(zipFile, ZipArchiveMode.Create))
                    {
                        for (int i = mail.Attachments.Count; i > 0; i--)
                        {
                            var attachment = mail.Attachments[i];
                            //var file = Path.GetTempFileName();
                            var file = Path.Combine(tempPath, Path.GetTempFileName());
                            attachment.SaveAsFile(file);
                            archive.CreateEntryFromFile(file, attachment.FileName);
                            File.Delete(file);
                            attachment.Delete();
                        }
                    }
                    //mail.Attachments.Remove();
                    mail.Attachments.Add(zipFile, OlAttachmentType.olByValue, 1, "Archive.zip");
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
