
namespace GlobalAddressListSync.Add_in
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.TabHome = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnSync = this.Factory.CreateRibbonButton();
            this.lblMessage = this.Factory.CreateRibbonLabel();
            this.lblOABLast = this.Factory.CreateRibbonLabel();
            this.TabHome.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabHome
            // 
            this.TabHome.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabHome.ControlId.OfficeId = "TabContacts";
            this.TabHome.Groups.Add(this.group1);
            this.TabHome.Label = "TabContacts";
            this.TabHome.Name = "TabHome";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnSync);
            this.group1.Items.Add(this.lblMessage);
            this.group1.Items.Add(this.lblOABLast);
            this.group1.Label = "GAL Sync";
            this.group1.Name = "group1";
            // 
            // btnSync
            // 
            this.btnSync.Label = "Sync";
            this.btnSync.Name = "btnSync";
            this.btnSync.OfficeImageId = "SynchronizeHtml";
            this.btnSync.ShowImage = true;
            this.btnSync.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSync_Click);
            // 
            // lblMessage
            // 
            this.lblMessage.Label = "label1";
            this.lblMessage.Name = "lblMessage";
            // 
            // lblOABLast
            // 
            this.lblOABLast.Label = "label1";
            this.lblOABLast.Name = "lblOABLast";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.TabHome);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.TabHome.ResumeLayout(false);
            this.TabHome.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabHome;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSync;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblMessage;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblOABLast;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
