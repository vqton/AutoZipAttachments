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
using Microsoft.Office.Interop.Outlook;

namespace AutoZipAttachments
{
    public interface IEmailFormatter
    {
        void FormatEmail(Outlook.MailItem mailItem);
    }
    public interface IEmailRecipientManager
    {
        void AddCC(Outlook.MailItem mailItem, string email);
        void AddBCC(Outlook.MailItem mailItem, string[] emails);
    }
    public interface IAttachmentCompressor
    {
        void CompressAttachments(Outlook.MailItem mailItem, string outputPath);
    }
    public class EmailFormatter : IEmailFormatter
    {
        public void FormatEmail(Outlook.MailItem mailItem)
        {
            if (mailItem == null) return;

            string plainBody = mailItem.Body;
            string formattedBody = $@"
                <p style=""font-family: Arial; line-height: 1.5; font-size: 14px;"">
                    {plainBody}
                </p>";
            mailItem.HTMLBody = formattedBody;
            mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            mailItem.HTMLBody = formattedBody;
        }
    }
    public class EmailRecipientManager : IEmailRecipientManager
    {
        public void AddCC(Outlook.MailItem mailItem, string email)
        {
            Outlook.Recipient recipient = mailItem.Recipients.Add(email);
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
    }
    public class AttachmentCompressor : IAttachmentCompressor
    {
        public void CompressAttachments(Outlook.MailItem mailItem, string outputPath)
        {
            const int sizeLimit = 5 * 1024 * 1024; // 5MB in bytes
            int totalSize = 0;

            // Calculate the total size of all attachments
            foreach (Outlook.Attachment attachment in mailItem.Attachments)
            {
                totalSize += attachment.Size;
            }

            // If the total size exceeds 5MB, compress non-image attachments
            if (totalSize > sizeLimit)
            {
                string tempPath = Path.GetTempPath();
                string zipFilePath = Path.Combine(outputPath, $"{GenerateSlug(mailItem.Subject)}.zip");

                using (var zipStream = new FileStream(zipFilePath, FileMode.Create))
                using (var zipWriter = WriterFactory.Open(zipStream, ArchiveType.Zip, CompressionType.Deflate))
                {
                    foreach (Outlook.Attachment attachment in mailItem.Attachments)
                    {
                        if (!IsImageFile(attachment.FileName))
                        {
                            string tempFileName = Path.Combine(tempPath, attachment.FileName);
                            attachment.SaveAsFile(tempFileName);

                            string sluggedFileName = GenerateSlug(Path.GetFileNameWithoutExtension(attachment.FileName)) + Path.GetExtension(attachment.FileName);
                            zipWriter.Write(sluggedFileName, tempFileName);

                            File.Delete(tempFileName);
                        }
                    }
                }

                // Remove non-image attachments and add the compressed zip file
                for (int i = mailItem.Attachments.Count; i > 0; i--)
                {
                    Outlook.Attachment attachment = mailItem.Attachments[i];
                    if (!IsImageFile(attachment.FileName))
                    {
                        attachment.Delete();
                    }
                }

                mailItem.Attachments.Add(zipFilePath, Outlook.OlAttachmentType.olByValue, Type.Missing, Path.GetFileName(zipFilePath));
            }
        }

        private bool IsImageFile(string fileName)
        {
            string[] imageExtensions = { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff" };
            string fileExtension = Path.GetExtension(fileName).ToLower();
            return imageExtensions.Contains(fileExtension);
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


    public class EmailSender
    {
        private readonly IEmailFormatter _emailFormatter;
        private readonly IEmailRecipientManager _emailRecipientManager;
        private readonly IAttachmentCompressor _attachmentCompressor;

        public EmailSender(IEmailFormatter emailFormatter, IEmailRecipientManager emailRecipientManager, IAttachmentCompressor attachmentCompressor)
        {
            _emailFormatter = emailFormatter;
            _emailRecipientManager = emailRecipientManager;
            _attachmentCompressor = attachmentCompressor;
        }

        public void FormatEmail(Outlook.MailItem mailItem)
        {
            _emailFormatter.FormatEmail(mailItem);
        }

        public void AddCC(Outlook.MailItem mailItem, string email)
        {
            _emailRecipientManager.AddCC(mailItem, email);
        }

        public void AddBCC(Outlook.MailItem mailItem, string[] emails)
        {
            _emailRecipientManager.AddBCC(mailItem, emails);
        }

        public void CompressAttachments(Outlook.MailItem mailItem, string outputPath)
        {
            _attachmentCompressor.CompressAttachments(mailItem, outputPath);
        }
    }

}





