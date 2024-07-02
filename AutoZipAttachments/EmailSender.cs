using System.IO;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;

using SharpCompress.Common;
using SharpCompress.Writers;
using System.Linq;

namespace AutoZipAttachments
{
    public interface IEmailSender
    {
        void AddCC(Outlook.MailItem mailItem);
        void AddBCC(Outlook.MailItem mailItem, string[] groupName);
        void CompressAttachments(Outlook.MailItem mailItem, string outputPath);
        void MoveTempDirectoryToSystemTemp(string sourcePath);

    }
    public class EmailSender : IEmailSender
    {



        public EmailSender()
        {

        }
        public void AddCC(Outlook.MailItem mailItem)
        {
            Outlook.Recipient recipient = mailItem.Recipients.Add("ton-vq@saigonco-op.com.vn");

            recipient.Type = (int)Outlook.OlMailRecipientType.olCC;
            recipient.Resolve();
        }

        public void AddBCC(Outlook.MailItem mailItem, string[] emails)
        {
            if (mailItem != null)
            {
                foreach (string email in emails)
                {
                    Outlook.Recipient bccRecipient = mailItem.Recipients.Add(email);
                    bccRecipient.Type = (int)Outlook.OlMailRecipientType.olBCC;
                    bccRecipient.Resolve();
                }
            }
        }
        public void MoveTempDirectoryToSystemTemp(string sourcePath)
        {
            string systemTempPath = Path.GetTempPath();
            string destPath = Path.Combine(systemTempPath, Path.GetFileName(sourcePath));

            if (Directory.Exists(sourcePath))
            {
                Directory.Move(sourcePath, destPath);
            }
        }
        public void CompressAttachments(Outlook.MailItem mailItem, string outputPath)
        {
            const long maxAttachmentSize = 5 * 1024 * 1024; // 5MB in bytes

            if (mailItem.Attachments.Count > 0)
            {
                long totalSize = 0;
                string[] excludedExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff" };

                // Calculate the total size of the attachments
                foreach (Outlook.Attachment attachment in mailItem.Attachments)
                {
                    string extension = Path.GetExtension(attachment.FileName).ToLower();
                    if (!excludedExtensions.Contains(extension))
                    {
                        totalSize += attachment.Size;
                    }
                }

                // Check if the total size exceeds the threshold
                if (totalSize > maxAttachmentSize)
                {
                    // Define the temporary directory on D: drive
                    string tempDirectory = Path.Combine(@"E:\temp", Guid.NewGuid().ToString());
                    Directory.CreateDirectory(tempDirectory);

                    try
                    {
                        // Save each attachment to the temporary directory, excluding image files
                        foreach (Outlook.Attachment attachment in mailItem.Attachments)
                        {
                            string attachmentPath = Path.Combine(tempDirectory, attachment.FileName);
                            string extension = Path.GetExtension(attachment.FileName).ToLower();

                            if (!excludedExtensions.Contains(extension))
                            {
                                attachment.SaveAsFile(attachmentPath);
                            }
                        }

                        // Compress the non-image attachments into a zip archive using SharpCompress
                        string archivePath = Path.Combine(outputPath, "attachments.zip");

                        using (var archiveStream = File.OpenWrite(archivePath))
                        using (var writer = WriterFactory.Open(archiveStream, ArchiveType.Zip, CompressionType.Deflate))
                        {
                            writer.WriteAll(tempDirectory, "*", SearchOption.AllDirectories);
                        }

                        // Verify that the compressed file was created and is not empty
                        FileInfo archiveFileInfo = new FileInfo(archivePath);
                        if (archiveFileInfo.Exists && archiveFileInfo.Length > 0)
                        {
                            // Remove original non-image attachments (iterate in reverse order)
                            for (int i = mailItem.Attachments.Count; i > 0; i--)
                            {
                                Outlook.Attachment attachment = mailItem.Attachments[i];
                                string extension = Path.GetExtension(attachment.FileName).ToLower();

                                if (!excludedExtensions.Contains(extension))
                                {
                                    attachment.Delete();
                                }
                            }

                            // Attach the compressed zip file
                            mailItem.Attachments.Add(archivePath, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                        }
                        else
                        {
                            throw new Exception("Compression failed or resulted in an empty archive.");
                        }
                    }
                    finally
                    {
                        // Cleanup temporary files
                        if (Directory.Exists(tempDirectory))
                        {
                            Directory.Delete(tempDirectory, recursive: true);
                        }
                    }
                }
            }
        }



    }



}
