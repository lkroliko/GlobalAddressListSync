# GlobalAddressListSync

If you want use GAL (Global Address List) in mobile phone and see who is calling, you need to copy all cotacts from GAL to own mailbox.
This Outlook add-in will do it for you and will synchronize when outlook see new GAL.

### Requirements
Outlook 2013 or newer.  
Tested with Outlook 2013.

### Installation
1) download file GlobalAddressListSync.zip from https://github.com/lkroliko/GlobalAddressListSync/releases
2) unzip
3) run setup and install

### How it works
Program will create new contact folder with name of GAL. For example it can be "Global Addess List" or in polish "Globalna lista adresów". 
It depends on your language. All contacts from GAL will be added to that folder. **NEVER USE THIS SYNC FOLDER TO STORE YOUR CONTACTS.**  
  
On Outlook startup add-in is checking if OAB is changed. If OAB is changed it meens GAL is changed and synchronization will start.
If contacts are deleted from GAL then it will be deleted from sync contacts folder. If contact is changed then will be updated. Email address is used as key to join GAL with folder.

You can force to sync. Go to contacts and find GAL Sync tab on ribbon then click "Sync".

### License
This project is licensed under the MIT License

### Help
Feel free to contact with me if you have problem or idea to improve this add-in.
