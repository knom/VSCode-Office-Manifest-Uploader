# Office Outlook MailApp Manifest Uploader
Uploads the [manifest.xml](https://msdn.microsoft.com/en-us/library/office/dn642483.aspx) of your [Outlook Mail App](https://msdn.microsoft.com/EN-US/library/office/fp161135.aspx) 
 into Office 365 or your Exchange Server.
 
## How to install
* Press (`Cmd+E` on OSX or `Ctrl+E` on Windows and Linux)
* Type `ext install office-mailapp-manifestuploader`
 and hit `enter`

## Usage
* Press `Cmd+Alt+U` on OSX or `Ctrl+Alt+U` on Windows and Linux
* Or press `F1` and type `Upload Outlook Mail App manifest`
* You will be prompted for username, password and server address of the target server.
* Optional: Configure these settings in your `User Settings` to not be prompted every time.

 
## Configuration options
User & server settings can be stored in `User Settings`.
| Option | Description |
|----------------------------------|--------------------------------------------------------------------------------------------------------------------------------------|
| officeManifestUploader.userName | The login for your Office 365 or Exchange, e.g. user@foo.com. |
| officeManifestUploader.password | The password. It's recommended NOT to store it, but rather skip this setting. Then you will be prompted for the password every time. |
| officeManifestUploader.serverUrl | The server address, e.g. mail.office365.com. |12345


##**Enjoy!**##