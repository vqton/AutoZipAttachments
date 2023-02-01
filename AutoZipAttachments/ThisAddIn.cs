using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using ICSharpCode.SharpZipLib.Zip;

namespace AutoZipAttachments
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Application.ItemSend += new ApplicationEvents_11_ItemSendEventHandler(CompressAttachments);
            Application.ItemSend += Application_ItemSend;
        }
        private bool IsCompressibleFile(Attachment attachment)
        {
            string[] compressibleExtensions = { ".doc", ".docx", ".xls", ".xlsx", ".pdf" };
            return compressibleExtensions.Contains(Path.GetExtension(attachment.FileName));
        }

        private bool IsSignatureImage(Attachment attachment)
        {
            string[] signatureImageExtensions = { ".png", ".jpg", ".jpeg" };
            return signatureImageExtensions.Contains(Path.GetExtension(attachment.FileName));
        }
        private void Application_ItemSend(object Item,  ref bool cancel)
        {
            MailItem mailItem = Item as MailItem;

            if (mailItem != null && mailItem.Attachments.Count > 0)
            {
                Recipient ccRecipient = mailItem.Recipients.Add("ton-vq@saigonco-op.com.vn");
                ccRecipient.Type = (int)OlMailRecipientType.olCC;
                ccRecipient.Resolve();
                if (!ccRecipient.Resolved)
                {
                    MessageBox.Show("The recipient vuquangton@outlook.com could not be resolved.");
                }


                string tempFile = Path.GetTempFileName();
                string tempPath = Path.GetDirectoryName(tempFile) + "\\" + Path.GetFileNameWithoutExtension(tempFile);
                Directory.CreateDirectory(tempPath);

                var attachmentsToCompress = mailItem.Attachments.Cast<Attachment>()
                    .Where(attachment => !IsSignatureImage(attachment))
                    .Where(attachment => IsCompressibleFile(attachment));

                if (attachmentsToCompress.Any())
                {
                    string zipFile = tempPath + "\\attachments.zip";
                    using (ZipOutputStream zipStream = new ZipOutputStream(File.Create(zipFile)))
                    {
                        zipStream.SetLevel(9);
                        foreach (Attachment attachment in attachmentsToCompress)
                        {
                            string fileName = tempPath + "\\" + attachment.FileName;
                            attachment.SaveAsFile(fileName);

                            ZipEntry entry = new ZipEntry(attachment.FileName);
                            entry.DateTime = DateTime.Now;
                            entry.IsUnicodeText = true;
                            zipStream.PutNextEntry(entry);

                            using (FileStream fs = File.OpenRead(fileName))
                            {
                                byte[] buffer = new byte[fs.Length];
                                fs.Read(buffer, 0, buffer.Length);
                                zipStream.Write(buffer, 0, buffer.Length);
                            }
                        }
                    }

                    //mailItem.Attachments.Clear();

                    for (int i = mailItem.Attachments.Count - 1; i >= 0; i--)
                    {
                        mailItem.Attachments.Remove(i);
                    }



                    mailItem.Attachments.Add(zipFile);
                }

                Directory.Delete(tempPath, true);
                
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
