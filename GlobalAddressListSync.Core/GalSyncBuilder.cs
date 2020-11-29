using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GlobalAddressListSync.Core
{
    public class GalSyncBuilder
    {
        private Application _application;

        public GalSyncBuilder SetApplication(Application application)
        {
            _application = application;
            return this;
        }

        public GalSync Build()
        {
            if (_application == null)
                throw new ArgumentNullException("application");
            return new GalSync(_application, new Folder(_application), new Options());
        }

        public static GalSyncBuilder New { get { return new GalSyncBuilder(); } }
    }
}
