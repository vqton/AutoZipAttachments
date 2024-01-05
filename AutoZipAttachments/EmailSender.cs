using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace AutoZipAttachments
{
    public interface IEmailSender
    {
        void AddCC(Outlook.MailItem mailItem);
        void AddBCC(Outlook.MailItem mailItem, string groupName);
    }
    public class EmailSender : IEmailSender
    {public interface IEmailSender
    {
        void AddCC(Outlook.MailItem mailItem);
        void AddBCC(Outlook.MailItem mailItem, string groupName);
    }
        public void AddCC(Outlook.MailItem mailItem)
        {
            Outlook.Recipient recipient = mailItem.Recipients.Add(mailItem.SenderEmailAddress);
            recipient.Type = (int)Outlook.OlMailRecipientType.olCC;
            recipient.Resolve();
        }

        public void AddBCC(Outlook.MailItem mailItem, string groupName)
        {
            Outlook.AddressEntry addressEntry = mailItem.Application.Session.CurrentUser.AddressEntry;
            if (addressEntry != null && addressEntry.Type == "EX")
            {
                Outlook.ExchangeUser manager = addressEntry.GetExchangeUser()?.GetExchangeUserManager();
                if (manager != null)
                {
                    Outlook.AddressEntries addrEntries = manager.GetDirectReports();
                    Outlook.DistListItem distList = null;

                    // Check if the group exists
                    foreach (Outlook.AddressEntry addrEntry in addrEntries)
                    {
                        if (addrEntry.Name == groupName)
                        {
                            distList = addrEntry.GetContact() as Outlook.DistListItem;
                            break;
                        }
                    }

                    // If the group does not exist, create it
                    if (distList == null)
                    {
                        distList = mailItem.Application.CreateItem(Outlook.OlItemType.olDistributionListItem) as Outlook.DistListItem;
                        if (distList != null)
                        {
                            distList.DLName = groupName;
                            Outlook.MailItem newMail = mailItem.Application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                            if (newMail != null)
                            {
                                Outlook.Recipient recipient = newMail.Recipients.Add("vuquangton@outlook.com");
                                distList.AddMembers(newMail.Recipients);
                                distList.Save();
                            }
                        }
                    }

                    // Add the group members to the BCC field
                    if (distList != null)
                    {
                        for (int i = 1; i <= distList.MemberCount; i++)
                        {
                            Outlook.Recipient memberRecipient = distList.GetMember(i);
                            if (memberRecipient != null)
                            {
                                Outlook.AddressEntry member = memberRecipient.AddressEntry;
                                if (member != null)
                                {
                                    Outlook.Recipient bccRecipient = mailItem.Recipients.Add(member.Address);
                                    bccRecipient.Type = (int)Outlook.OlMailRecipientType.olBCC;
                                    bccRecipient.Resolve();
                                }
                            }
                        }
                    }
                }
            }
        }








    }
}
