using Microsoft.Office.Interop.Outlook;

namespace GlobalAddressListSync.Core
{
    public static class ContactItemExtensions
    {
        public static void Update(this ContactItem thisContactItem, ExchangeUser user)
        {
            thisContactItem.Email1Address = user.PrimarySmtpAddress;
            thisContactItem.BusinessTelephoneNumber = user.BusinessTelephoneNumber;
            thisContactItem.MobileTelephoneNumber = user.MobileTelephoneNumber;
            thisContactItem.CompanyName = user.CompanyName;
            thisContactItem.Department = user.Department;
            thisContactItem.FirstName = user.FirstName;
            thisContactItem.LastName = user.LastName;
            thisContactItem.FullName = user.Name;
            thisContactItem.JobTitle = user.JobTitle;
            thisContactItem.OfficeLocation = user.OfficeLocation;

            thisContactItem.Save();
        }

        public static void Update(this ContactItem thisContactItem, ExchangeDistributionList distributionList)
        {
            var user = distributionList.GetExchangeUser();
            if (user != null)
            {
                thisContactItem.Email1Address = distributionList.PrimarySmtpAddress;
                thisContactItem.LastName = distributionList.Name;
                thisContactItem.FirstName = string.Empty;
                thisContactItem.BusinessTelephoneNumber = user.BusinessTelephoneNumber;
                thisContactItem.MobileTelephoneNumber = user.MobileTelephoneNumber;
                thisContactItem.CompanyName = user.CompanyName;
                thisContactItem.Department = user.Department;
                thisContactItem.OfficeLocation = user.OfficeLocation;
                thisContactItem.Save();
            }
        }
    }
}
