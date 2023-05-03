# O365-Email-Attachment-Extractor
This repository contains a Python script that can be used to extract email attachments from Office 365. 

I needed a script that could extract PDF attachments of invoices I receive. I found bits and pieces there that didn't work for me hence I decided to bring those bits and pieces together.


To use the script, you will need to create an Azure AD application and grant it permissions to access the Microsoft Graph API. You will also need to install the required Python libraries, such as requests and msal.  This script can be useful for automating the extraction of email attachments from Office 365, for example for archiving or backup purposes. Feel free to use, modify, and contribute to this script as needed.


The script uses the Microsoft Graph API to authenticate with Office 365 and retrieve the specified email messages, and then saves the attachments to a local directory.  

The script can be customized to extract attachments based on various criteria, such as the sender, recipient, subject, or date range of the email message. It can also be adapted to filter with  different types of attachments, including documents, images, and zip files.  


In your case depending on your level of access to Azure AD, you may want to use me Graph endpoint instead of users.


