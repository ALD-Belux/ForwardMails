namespace ForwardMailsService
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.ForwardMailSVCProcessInstaller1 = new System.ServiceProcess.ServiceProcessInstaller();
            this.ForwardMailSVCInstaller1 = new System.ServiceProcess.ServiceInstaller();
            // 
            // ForwardMailSVCProcessInstaller1
            // 
            this.ForwardMailSVCProcessInstaller1.Password = null;
            this.ForwardMailSVCProcessInstaller1.Username = null;
            // 
            // ForwardMailSVCInstaller1
            // 
            this.ForwardMailSVCInstaller1.Description = "Forward mail from a specific mailbox folder to an address";
            this.ForwardMailSVCInstaller1.DisplayName = "Mails Forwarder";
            this.ForwardMailSVCInstaller1.ServiceName = "FWMails";
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.ForwardMailSVCProcessInstaller1,
            this.ForwardMailSVCInstaller1});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller ForwardMailSVCProcessInstaller1;
        private System.ServiceProcess.ServiceInstaller ForwardMailSVCInstaller1;
    }
}