namespace ConsultasWebExcelAddin
{
    partial class RibbonMain : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMain()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Designer de Componentes

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabConsultasWs = this.Factory.CreateRibbonTab();
            this.groupReceitaCnpj = this.Factory.CreateRibbonGroup();
            this.btnConsultaCNPJSCelulas = this.Factory.CreateRibbonButton();
            this.groupLogistica = this.Factory.CreateRibbonGroup();
            this.btnBuscarCEPCelulas = this.Factory.CreateRibbonButton();
            this.groupHelp = this.Factory.CreateRibbonGroup();
            this.btnHelp = this.Factory.CreateRibbonButton();
            this.tabConsultasWs.SuspendLayout();
            this.groupReceitaCnpj.SuspendLayout();
            this.groupLogistica.SuspendLayout();
            this.groupHelp.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabConsultasWs
            // 
            this.tabConsultasWs.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabConsultasWs.Groups.Add(this.groupReceitaCnpj);
            this.tabConsultasWs.Groups.Add(this.groupLogistica);
            this.tabConsultasWs.Groups.Add(this.groupHelp);
            this.tabConsultasWs.Label = "Consultas Online";
            this.tabConsultasWs.Name = "tabConsultasWs";
            // 
            // groupReceitaCnpj
            // 
            this.groupReceitaCnpj.Items.Add(this.btnConsultaCNPJSCelulas);
            this.groupReceitaCnpj.Label = "CNPJ";
            this.groupReceitaCnpj.Name = "groupReceitaCnpj";
            // 
            // btnConsultaCNPJSCelulas
            // 
            this.btnConsultaCNPJSCelulas.Label = "Buscar CNPJs";
            this.btnConsultaCNPJSCelulas.Name = "btnConsultaCNPJSCelulas";
            this.btnConsultaCNPJSCelulas.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConsultaCNPJSCelulas_Click);
            // 
            // groupLogistica
            // 
            this.groupLogistica.Items.Add(this.btnBuscarCEPCelulas);
            this.groupLogistica.Label = "Correios";
            this.groupLogistica.Name = "groupLogistica";
            // 
            // btnBuscarCEPCelulas
            // 
            this.btnBuscarCEPCelulas.Label = "Buscar CEPs";
            this.btnBuscarCEPCelulas.Name = "btnBuscarCEPCelulas";
            this.btnBuscarCEPCelulas.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBuscarCEPCelulas_Click);
            // 
            // groupHelp
            // 
            this.groupHelp.Items.Add(this.btnHelp);
            this.groupHelp.Label = "Ajuda";
            this.groupHelp.Name = "groupHelp";
            // 
            // btnHelp
            // 
            this.btnHelp.Label = "Ajuda";
            this.btnHelp.Name = "btnHelp";
            // 
            // RibbonMain
            // 
            this.Name = "RibbonMain";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabConsultasWs);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonMain_Load);
            this.tabConsultasWs.ResumeLayout(false);
            this.tabConsultasWs.PerformLayout();
            this.groupReceitaCnpj.ResumeLayout(false);
            this.groupReceitaCnpj.PerformLayout();
            this.groupLogistica.ResumeLayout(false);
            this.groupLogistica.PerformLayout();
            this.groupHelp.ResumeLayout(false);
            this.groupHelp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabConsultasWs;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupReceitaCnpj;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupLogistica;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConsultaCNPJSCelulas;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBuscarCEPCelulas;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelp;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonMain RibbonMain
        {
            get { return this.GetRibbon<RibbonMain>(); }
        }
    }
}
