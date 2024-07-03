using System.IO;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Unidecode.NET;
using SharpCompress.Common;
using SharpCompress.Writers;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

namespace AutoZipAttachments
{
    public interface IEmailSender
    {
        void AddCC(Outlook.MailItem mailItem);
        void AddBCC(Outlook.MailItem mailItem, string[] groupName);
        void CompressAttachments(Outlook.MailItem mailItem, string outputPath);
        void FormatEmail(Outlook.MailItem mailItem);

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


        public void FormatEmail(Outlook.MailItem mailItem)
        {
            return;
            // Check if the mail item is not null
            if (mailItem == null) return;

            // Step 1: Convert the mail body to plain text
            string plainBody = mailItem.Body;

            // Step 2: Convert the plain text to HTML and apply the styles
            string formattedBody = $@"
                    <p style=""font-family: Arial; line-height: 1.5; font-size: 14px;"">
                        {plainBody}
                    </p>";
            mailItem.HTMLBody = formattedBody;

            // Set the mail item body format to HTML and assign the formatted body
            mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            mailItem.HTMLBody = formattedBody;
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
                                // Ensure the filename is saved in Unicode
                                attachment.SaveAsFile(attachmentPath);
                            }
                        }

                        // Generate slug from the email subject
                        string slug = GenerateSlug(mailItem.Subject);

                        // Compress the non-image attachments into a zip archive using SharpCompress
                        string archivePath = Path.Combine(outputPath, $"{slug}.zip");

                        using (var archiveStream = File.OpenWrite(archivePath))
                        using (var writer = WriterFactory.Open(archiveStream, ArchiveType.Zip, CompressionType.Deflate))
                        {
                            var files = Directory.EnumerateFiles(tempDirectory, "*", SearchOption.AllDirectories);
                            foreach (var file in files)
                            {
                                var entryName = file.Substring(tempDirectory.Length + 1);
                                writer.Write(entryName, file);
                            }
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

        private string GenerateSlug(string title)
        {

            if (string.IsNullOrWhiteSpace(title))
            {
                return "attachments";
            }

            // Normalize the string to remove diacritics and convert to ASCII
            string normalizedTitle = title.Unidecode();

            // Use StringBuilder to build the slug
            StringBuilder slugBuilder = new StringBuilder();
            foreach (char c in normalizedTitle)
            {
                // Check if the character is a letter or digit
                if (char.IsLetterOrDigit(c))
                {
                    slugBuilder.Append(char.ToLowerInvariant(c));
                }
                // Replace spaces with hyphens
                else if (char.IsWhiteSpace(c))
                {
                    slugBuilder.Append('-');
                }
                // Replace non-letter, non-digit characters with nothing (remove them)
                else if (c == '-' || c == '_')
                {
                    slugBuilder.Append(c);
                }
            }

            // Remove multiple consecutive hyphens
            string slug = Regex.Replace(slugBuilder.ToString(), "-{2,}", "-").Trim('-');

            // Ensure the slug is not empty
            return string.IsNullOrWhiteSpace(slug) ? "attachments" : slug;
        }
    }




}
