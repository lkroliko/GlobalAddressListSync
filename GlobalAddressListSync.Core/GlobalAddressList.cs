using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GlobalAddressListSync.Core
{
    public class GlobalAddressList
    {
        private Application _application { get; }

        public GlobalAddressList(Outlook.Application application)
        {
            _application = application;
        }


        private void a()
        {

            Outlook.AddressList gal = _application.Session.GetGlobalAddressList();
            foreach (Outlook.AddressEntry address in gal.AddressEntries)
            {
                
                //if (address.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry)
                //{
                //    address
                //}
            }
        }
    }
}
