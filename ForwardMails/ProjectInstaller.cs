using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Threading.Tasks;

namespace ForwardMailsService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }

        public override void Install(System.Collections.IDictionary stateSaver)
        {
            RetrieveServiceName();
            base.Install(stateSaver);
        }

        public override void Uninstall(System.Collections.IDictionary savedState)
        {
            RetrieveServiceName();
            base.Uninstall(savedState);
        }

        private void RetrieveServiceName()
        {
            var serviceName = Context.Parameters["servicename"];
            var displayName = Context.Parameters["displayname"];
            if (!string.IsNullOrEmpty(serviceName))
            {
                this.ForwardMailSVCInstaller1.ServiceName = serviceName;
                this.ForwardMailSVCInstaller1.DisplayName = displayName;
            }
        }
    }
}
