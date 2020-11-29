using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GlobalAddressListSync.Core
{
    public class Folder
    {
        private Application _application { get; }
        private string _folderName;

        public Folder(Outlook.Application application)
        {
            _application = application;
            _folderName = application.Session.GetGlobalAddressList().Name;
        }

        public Outlook.MAPIFolder GetSyncFolder()
        {
            if (FolderExists(out Outlook.Folder folder) == false)
                return CreateContactsFolder();
            return folder;
        }

        private Outlook.MAPIFolder CreateContactsFolder()
        {
            var contacts = GetDefaultContactsFolder();
            var parent = contacts.Parent as Outlook.MAPIFolder;
            return parent.Folders.Add(_folderName, Outlook.OlDefaultFolders.olFolderContacts);
        }

        private bool FolderExists(out Outlook.Folder folder)
        {
            var contacts = GetDefaultContactsFolder();
            var parent = contacts.Parent as Outlook.Folder;
            foreach (var item in parent.Folders)
                if (item is Outlook.Folder && (item as Outlook.Folder).Name == _folderName)
                {
                    folder = item as Outlook.Folder;
                    return true;
                }
            folder = null;
            return false;
        }

        public Outlook.Folder GetDefaultContactsFolder()
        {
            return (Outlook.Folder)_application.Session.GetDefaultFolder
                   (Outlook.OlDefaultFolders.olFolderContacts);
        }

        public List<ContactItem> GetContactsList()
        {
            var list = new List<ContactItem>();
            var folder = GetSyncFolder();
            foreach (var item in folder.Items)
                if (item is ContactItem)
                    list.Add(item as ContactItem);
            return list;
        }

        public void AddExchangeUser(ExchangeUser user)
        {
            ContactItem newContact = GetSyncFolder().Items.Add("IPM.Contact") as ContactItem;
            newContact.Update(user);
        }

        public void AddExchangeDistributionList(ExchangeDistributionList distributionList)
        {
            ContactItem newContact = GetSyncFolder().Items.Add("IPM.Contact") as ContactItem;
            newContact.Update(distributionList);
        }
    }
}
