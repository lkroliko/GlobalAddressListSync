using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GlobalAddressListSync.Core
{
    public class Options
    {
        private string _galKey = "SOFTWARE\\Microsoft\\Exchange\\Exchange Provider\\OABs";
        private const string _oabKey = "OAB Last Modified Time";
        private const string _galsKey = "GALS Last Sync Time";

        public DateTime? OABLastModifiedTime { get { return GetDate(_oabKey); } }
        public DateTime? GALSLastSyncTime { get { return GetDate(_galsKey); } set { SetDate(_galsKey, value); } }

        public Options()
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(_galKey);

            var subKey = key.GetSubKeyNames();
            if (subKey.Length > 0)
                _galKey = $"{_galKey}\\{subKey[0]}";
        }

        private DateTime? GetDate(string key)
        {
            RegistryKey regKey = Registry.CurrentUser.OpenSubKey(_galKey);
            byte[] value = (byte[])regKey.GetValue(key);
            if (value != null)
                return DateTime.FromFileTime(BitConverter.ToInt64(value, 0));
            return null;
        }

        public void SetDate(string key, DateTime? dateTime)
        {
            RegistryKey regKey = Registry.CurrentUser.OpenSubKey(_galKey, true);
            if (dateTime.HasValue)
            {
                long value = dateTime.Value.ToFileTime();
                regKey.SetValue(key, BitConverter.GetBytes(value), RegistryValueKind.Binary);
            }
            else
            {
                regKey.DeleteSubKey(_galsKey);
            }
        }
    }
}
