wa2mc
=====

Wild Apricot to MailChimp

MailChimp provides a free service for up to 2000 subscribers and up to 12000 mails per month. They have a powerful web-based interface to create emails that are optimized for mobile devices. Wild Apricot limits the number of contacts and members to 500 with their basic paid account. Adding a lot of email list subscribers to Wild Apricot may require an update of the Wild Apricot package.

wa2mc.exe subscribes contacts and members from Wild Apricot to a MailChimp list. Best practice for MailChimp is using one list with groups. The application adds a status for members and contacts. The status will be blank for those who subscribe from the MailChimp form. Members and contacts that have been added to the list will get a message that they are already signed up when they enter their email address into the web signup form.

**Usage:**

* Download ZIP file.
* Edit  the wa2mc.exe.config in any editor. Edit your values for the MailChimpListID, MailChimpGroupID, MailChimpKey and WildApricotKey. If the eMail addresses are saved in another field than “E-Mail”, enter the field name to EMailField. You can enter a value to EMailField2 to use this field if the eMail address in the first field is empty or not valid.
* Run the application local and check the MailChimp list. It may take a couple of minutes until the added addresses are a available in MailChimp.
* You may change the Crypt key in wa2mc.exe.config to True and restart the application. The IDs and keys will be saved encrypted.
* You can schedule the application local. You can also use a Windows Azure Webjob. Zip the folder after you tested locally and encrypted the configuration file and upload it as a webjob.

Requirements:

Wild Apricot API must be enabled.
The search string used for the Wild Apricot API is 'Email delivery disabled' eq false AND Archived eq false AND 'Member emails and newsletters' eq true to be compliant with MailChimp.
A MailChimp API key is required.
A MailChimp list with a Status group and options Member and Contact. You can get the MailChimpGroupID from the source code of the MailChimp webpage that displays the group button.
Limitations:

Members and contacts that update the email address in Wild Apricot will be added to MailChimp using the new email address. They will receive MailChimp newsletters to both email address until they unsubscribe with the old email address.
