using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GlobalAddressListSync.Core
{
    public class GalSync
    {
        private Application _application;
        private Folder _folder;
        public Options Options { get; }
        private bool _processExchangeUsers = true;
        private bool _processExchangeDistributionLists = false;
        private List<ContactItem> _contacts;
        private AddressList _globalAddressList;

        internal GalSync(Outlook.Application application, Folder folder, Options options)
        {
            _application = application;
            _folder = folder;
            Options = options;
        }

        public void Sync()
        {
            if (Options.GALSLastSyncTime == null || Options.GALSLastSyncTime < Options.OABLastModifiedTime)
                SyncForced();
        }

        public void SyncForced()
        {
            _contacts = _folder.GetContactsList();
           
            _globalAddressList = _application.Session.GetGlobalAddressList();

            for (int i = 1; i < _globalAddressList.AddressEntries.Count - 1; i++)
                ProcessAddressEntry(_globalAddressList.AddressEntries[i]);

            _contacts.ForEach(c => c.Delete());
            _contacts.Clear();
            Options.GALSLastSyncTime = DateTime.Now;
        }

        private void ProcessAddressEntry(AddressEntry addressEntry)
        {
            if (_processExchangeUsers)
            {
                bool isExchangeUser = addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry
                                    || addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry;
                if (isExchangeUser)
                    ProcessExchangeUser(addressEntry.GetExchangeUser());
            }

            if (_processExchangeDistributionLists)
            {
                bool isExchangeDistribution = addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeDistributionListAddressEntry;
                if (isExchangeDistribution)
                    ProcessExchangeDistributionList(addressEntry.GetExchangeDistributionList());
            }
        }

        private void ProcessExchangeUser(ExchangeUser user)
        {
            if (user == null)
                return;
            var contact = _contacts.FirstOrDefault(c => c.Email1Address == user.PrimarySmtpAddress);

            if (contact == null)
            {
                _folder.AddExchangeUser(user);
            }
            else
            {
                contact.Update(user);
                _contacts.Remove(contact);
            }
        }

        private void ProcessExchangeDistributionList(ExchangeDistributionList distributionList)
        {
            if (distributionList == null)
                return;
            var contact = _contacts.FirstOrDefault(c => c.Email1Address == distributionList.PrimarySmtpAddress);

            if (contact == null)
            {
                _folder.AddExchangeDistributionList(distributionList);
            }
            else
            {
                contact.Update(distributionList);
                _contacts.Remove(contact);
            }
        }
    }
}
