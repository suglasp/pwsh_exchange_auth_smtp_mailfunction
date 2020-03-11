
Example code to send a authenticated mail through an on-premise Exchange server from Powershell.

There are two scripts : 
- pwsh_exch_smtp_gencred.ps1 : generates a hashed clixml file that holds the credential file for the SMTP authentication.
- pwsh_exch_smtp_send.ps1    : example function for sending an authenticated mail, loading the clixml file.

The function can include attachments.

Other files:
- smtp_addresses.txt : a text file that holds mail adresses where the function needs to send to (recipients)
- test.txt : a demo file to show the function can send attachments.
