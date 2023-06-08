# AutoZipAttachments
Small Auto Zip Attachments Outlook Add-in
# AutoZipAttachment Outlook Add-in
This Outlook add-in is built in C# and automatically compresses file attachments, copies the sender, and BCCs an email address for backup purposes.
## Key Features
### 1. Auto compress files attachment. 
This add-in will automatically compress any file attachments into a ZIP file before sending the email. This helps reduce the size of emails and ensures attachments are sent in a compressed, optimized format.
### 2. Always cc to sender (note your self). 
This add-in will automatically add the sender's email address in the CC field of any outgoing email. This helps the sender keep a copy of the sent email in their Sent folder for their records.
### 3. Always bcc to email address in order to backup.
This add-in will automatically add an email address in the BCC field for backup purposes. All outgoing emails will be BCC'd to this address.
### 4. The add in build in C#.
This Outlook Add-in is built using C# and the Office SDK. The source code is included in this repository.
## How to Install
1. Download the AutoZipAttachment.outlookaddin file. 
2. In Outlook, go to File > Options > Add-Ins. 
3. Click "Go..." next to "Manage COM Add-ins". 
4. Click "Add..." and browse to select the AutoZipAttachment.outlookaddin file you downloaded.
5. Check the "AutoZipAttachment" add-in in the list and click "OK". 
6. The AutoZipAttachment add-in is now installed and will automatically run when you compose new emails in Outlook.
