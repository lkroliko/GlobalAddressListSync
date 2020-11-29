using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GlobalAddressListSync.Core
{
    public class UserAddressBook
    {
        private Application _application { get; }

        public UserAddressBook(Outlook.Application application)
        {
            _application = application;
        }

        public IEnumerable<Outlook.ContactItem> GetContacts()
        {
            Outlook.MAPIFolder contacts = (Outlook.MAPIFolder)_application.ActiveExplorer().Session.GetDefaultFolder
                   (Outlook.OlDefaultFolders.olFolderContacts);
            foreach (var contact in contacts.Items)
                yield return contact as Outlook.ContactItem;
        }
    }
}
