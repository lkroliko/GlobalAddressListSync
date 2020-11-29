using GlobalAddressListSync.Core;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GlobalAddressListSync.Add_in
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            var options = new Options();
            lblMessage.Label = $"Last sync {options.GALSLastSyncTime}";
        }

        private async void btnSync_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Outlook.Application application = Globals.ThisAddIn.Application;
                btnSync.Enabled = false;
                var galSync = GalSyncBuilder.New.SetApplication(application).Build();
                lblMessage.Label = $"Sync in progress";
                await Task.Run(() => galSync.SyncForced());
                btnSync.Enabled = true;
                lblMessage.Label = $"Last sync {galSync.Options.GALSLastSyncTime}";
            }
            catch
            {
                lblMessage.Label = "Error";
            }
        }
    }
}
