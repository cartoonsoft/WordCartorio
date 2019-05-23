namespace AdmCartorioWord
{
    partial class FaixaOpcoes : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public FaixaOpcoes()
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
            this.GerenciarCamposGroup = this.Factory.CreateRibbonGroup();
            this.CbNatureza = this.Factory.CreateRibbonComboBox();
            this.CbCampo = this.Factory.CreateRibbonComboBox();
            this.BtnAdicionar = this.Factory.CreateRibbonButton();
            this.GerenciadorRepeticaoGroup = this.Factory.CreateRibbonGroup();
            this.btnOutorgantes = this.Factory.CreateRibbonButton();
            this.btnOutorgados = this.Factory.CreateRibbonButton();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.GerenciarCamposGroup.SuspendLayout();
            this.GerenciadorRepeticaoGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.GerenciarCamposGroup);
            this.tab1.Groups.Add(this.GerenciadorRepeticaoGroup);
            this.tab1.Label = "Adm Cartório";
            this.tab1.Name = "tab1";
            // 
            // GerenciarCamposGroup
            // 
            this.GerenciarCamposGroup.Items.Add(this.CbNatureza);
            this.GerenciarCamposGroup.Items.Add(this.CbCampo);
            this.GerenciarCamposGroup.Items.Add(this.BtnAdicionar);
            this.GerenciarCamposGroup.Label = "Gerenciador de campos";
            this.GerenciarCamposGroup.Name = "GerenciarCamposGroup";
            // 
            // CbNatureza
            // 
            this.CbNatureza.Label = "Natureza";
            this.CbNatureza.MaxLength = 100;
            this.CbNatureza.Name = "CbNatureza";
            this.CbNatureza.Text = "Selecione";
            this.CbNatureza.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CbNatureza_TextChanged);
            // 
            // CbCampo
            // 
            this.CbCampo.Label = "Campo";
            this.CbCampo.Name = "CbCampo";
            this.CbCampo.Text = null;
            this.CbCampo.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CbCampo_TextChanged);
            // 
            // BtnAdicionar
            // 
            this.BtnAdicionar.Label = "Adicionar";
            this.BtnAdicionar.Name = "BtnAdicionar";
            this.BtnAdicionar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAdicionar_Click);
            // 
            // GerenciadorRepeticaoGroup
            // 
            this.GerenciadorRepeticaoGroup.Items.Add(this.label1);
            this.GerenciadorRepeticaoGroup.Items.Add(this.btnOutorgantes);
            this.GerenciadorRepeticaoGroup.Items.Add(this.btnOutorgados);
            this.GerenciadorRepeticaoGroup.Label = "Repetições";
            this.GerenciadorRepeticaoGroup.Name = "GerenciadorRepeticaoGroup";
            // 
            // btnOutorgantes
            // 
            this.btnOutorgantes.Label = "Outorgantes";
            this.btnOutorgantes.Name = "btnOutorgantes";
            this.btnOutorgantes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnOutorgantes_Click);
            // 
            // btnOutorgados
            // 
            this.btnOutorgados.Label = "Outorgados";
            this.btnOutorgados.Name = "btnOutorgados";
            this.btnOutorgados.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnOutorgados_Click);
            // 
            // label1
            // 
            this.label1.Label = "Selecione oque deseja repetir";
            this.label1.Name = "label1";
            // 
            // FaixaOpcoes
            // 
            this.Name = "FaixaOpcoes";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.FaixaOpcoes_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.GerenciarCamposGroup.ResumeLayout(false);
            this.GerenciarCamposGroup.PerformLayout();
            this.GerenciadorRepeticaoGroup.ResumeLayout(false);
            this.GerenciadorRepeticaoGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GerenciarCamposGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox CbNatureza;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox CbCampo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAdicionar;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GerenciadorRepeticaoGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOutorgantes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOutorgados;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
    }

    partial class ThisRibbonCollection
    {
        internal FaixaOpcoes FaixaOpcoes
        {
            get { return this.GetRibbon<FaixaOpcoes>(); }
        }
    }
}
