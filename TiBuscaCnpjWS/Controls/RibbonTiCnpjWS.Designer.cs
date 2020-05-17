namespace TiBuscaCnpjWS
{
    partial class RibbonTiCnpjWS : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonTiCnpjWS()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Buscas = this.Factory.CreateRibbonGroup();
            this.BtnCallWsIntoPlan = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Buscas.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Buscas);
            this.tab1.Label = "Busca CNPJs";
            this.tab1.Name = "tab1";
            // 
            // Buscas
            // 
            this.Buscas.Items.Add(this.BtnCallWsIntoPlan);
            this.Buscas.Label = "Buscar CNPJ";
            this.Buscas.Name = "Buscas";
            // 
            // BtnCallWsIntoPlan
            // 
            this.BtnCallWsIntoPlan.Image = global::TiBuscaCnpjWS.Properties.Resources.media_loop_square_logo;
            this.BtnCallWsIntoPlan.Label = "Buscar CNPJs Selecionados";
            this.BtnCallWsIntoPlan.Name = "BtnCallWsIntoPlan";
            this.BtnCallWsIntoPlan.ShowImage = true;
            this.BtnCallWsIntoPlan.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCallWsIntoPlan_Click);
            // 
            // RibbonTiCnpjWS
            // 
            this.Name = "RibbonTiCnpjWS";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Buscas.ResumeLayout(false);
            this.Buscas.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Buscas;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnCallWsIntoPlan;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonTiCnpjWS Ribbon1
        {
            get { return this.GetRibbon<RibbonTiCnpjWS>(); }
        }
    }
}
