using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using ConsultasWebExcelAddin.wsCorreios;
using ConsultasWebExcelAddin.WebService;
using Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows;

namespace ConsultasWebExcelAddin
{
    public partial class RibbonMain
    {
        Worksheet currentSheet;
        Range currentCell;
        Range currentSelection;
        
        private void setWorkingRanges()
        {
            this.currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            this.currentCell = Globals.ThisAddIn.getCurrentCell();
            this.currentSelection = Globals.ThisAddIn.getSelectedRange();
        }

        private void RibbonMain_Load(object sender, RibbonUIEventArgs e)
        {
            setWorkingRanges();
        }

        private void btnBuscarCEPCelulas_Click(object sender, RibbonControlEventArgs e)
        {

            setWorkingRanges();
            showLoading();
            Task.Run(() => BuscaCEPFromCorreioWs());
        }

        private void btnConsultaCNPJSCelulas_Click(object sender, RibbonControlEventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Está operação preenche as 36 colunas imediatamente depois dos CNPJs selecionados, deseja continuar?",
                                                        "Está operação não é reversível",
                                                        MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.No)
            {
                return;
            }


            showLoading();
            setWorkingRanges();
            Task.Run(() => BuscaCnpjFromWs());
        }

        private void showLoading(bool showHide = true)
        {
            //@TODO: disable individual buttons instead of hidding them
            groupHelp.Visible = !showHide;
            groupLogistica.Visible = !showHide;
            groupReceitaCnpj.Visible = !showHide;
            groupLoading.Visible = showHide;
        }

        private void BuscaCnpjFromWs(bool areaDeTransferencia = false)
        {
            try
            {
                List<dynamic> DadosCnpj;

                Worksheet currentSheet = this.currentSheet;
                Range currentCell = this.currentCell;
                Range currentSelection = this.currentSelection;
                List<string> cnpjs = new List<string>();

                foreach (Range cellCnpj in currentSelection)
                {
                    if (cellCnpj.Value2 != null)
                    {
                       cnpjs.Add(cellCnpj.Value2.ToString());
                    }
                }

                DadosCnpj = TiCnpjConsumer.getFullDataByCnpj(cnpjs);

                if (areaDeTransferencia)
                {
                    toClipboard(DadosCnpj);
                }
                else
                {
                    toCurrentPlan(DadosCnpj);
                };
            }
            catch (System.Net.WebException ex)
            {
                MessageBox.Show("Verifique sua conexão com a internet" + ex.Message,
                                "Consulta CNPJ - Erro de Rede!",
                                MessageBoxButtons.OK);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Ocorreu um erro ao buscar informação do CNPJ", MessageBoxButtons.OK);
            }
            finally
            {
                showLoading(false);
            }
        }


        private void BuscaCEPFromCorreioWs(bool areaDeTransferencia = false)
        {
            try
            {

                Worksheet currentSheet = this.currentSheet;
                Range currentCell = this.currentCell;
                Range currentSelection = this.currentSelection;

                List<enderecoERP> ceps = new List<enderecoERP>();

                foreach (Range cellCnpj in currentSelection)
                {
                    if (cellCnpj.Value2 != null)
                    {
                        ceps.Add( CorreiosConsumer.getFullAddressFromCorreios(cellCnpj.Value2.ToString()));
                    }
                }

                if(areaDeTransferencia)
                {
                    toClipboard(ceps);   
                } 
                else
                {
                    toCurrentPlan(ceps);
                }

            }
            catch (System.Net.WebException ex)
            {
                MessageBox.Show("Verifique sua conexão com a internet" + ex.Message,
                                "Consulta CNPJ - Erro de Rede!",
                                MessageBoxButtons.OK);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Ocorreu um erro ao buscar informação do CEP: " + ex.Message, MessageBoxButtons.OK);
            }
            finally
            {
                showLoading(false);
            }
        }

        private void toClipboard(List<enderecoERP> ceps)
        {
            string clipboard = "";
            foreach (enderecoERP dadosCep in ceps)
            {    
                clipboard += dadosCep.end;
                clipboard += ";"+dadosCep.bairro;
                clipboard += ";"+dadosCep.cidade;
                clipboard += ";"+dadosCep.uf;
                clipboard += ";"+dadosCep.unidadesPostagem+"\r\n";
            }

            System.Windows.Clipboard.SetText(clipboard);
        }

        public void toClipboard(List<dynamic> dadosCnpj)
        {
            string clipboard = "";
            foreach (dynamic dados in dadosCnpj)
            {
                clipboard += dados.matriz_filial;
                clipboard += ";"+dados.razao_social;
                clipboard += ";"+dados.nome_fantasia;
                clipboard += ";"+dados.situacao;
                clipboard += ";"+dados.data_situacao;
                clipboard += ";"+dados.motivo_situacao;
                clipboard += ";"+dados.nm_cidade_exterior;
                clipboard += ";"+dados.cod_pais;
                clipboard += ";"+dados.nome_pais;
                clipboard += ";"+dados.cod_nat_juridica;
                clipboard += ";"+dados.data_inicio_ativ;
                clipboard += ";"+dados.cnae_fiscal;
                clipboard += ";"+dados.tipo_logradouro;
                clipboard += ";"+dados.logradouro;
                clipboard += ";"+dados.numero;
                clipboard += ";"+dados.complemento;
                clipboard += ";"+dados.bairro;
                clipboard += ";"+dados.cep;
                clipboard += ";"+dados.uf;
                clipboard += ";"+dados.municipio;
                clipboard += ";"+dados.ddd_1;
                clipboard += ";"+dados.telefone_1;
                clipboard += ";"+dados.ddd_2;
                clipboard += ";"+dados.telefone_2;
                clipboard += ";"+dados.ddd_fax;
                clipboard += ";"+dados.num_fax;
                clipboard += ";"+dados.email;
                clipboard += ";"+dados.qualif_resp;
                clipboard += ";"+dados.porte;
                clipboard += ";"+dados.opc_simples;
                clipboard += ";"+dados.data_opc_simples;
                clipboard += ";"+dados.data_exc_simples;
                clipboard += ";"+dados.opc_mei;
                clipboard += ";"+dados.sit_especial;
                clipboard += ";"+dados.data_sit_especial;
                clipboard += ";" + dados.capital_social + "\r\n";

                System.Windows.Clipboard.SetText(clipboard);
            }

        }

        public void toCurrentPlan(List<enderecoERP> enderecos)
        {
            int currentRow = currentSelection.Cells[1, 1].row; //use this instead of currentCell to avoid misspositioning
            foreach (enderecoERP dadosCep in enderecos)
            {
                int currentColumn = currentSelection.Cells[1,1].Column; //use this instead of currentCell to avoid misspositioning
                currentSheet.Cells[currentRow, ++currentColumn] = dadosCep.end;
                currentSheet.Cells[currentRow, ++currentColumn] = dadosCep.bairro;
                currentSheet.Cells[currentRow, ++currentColumn] = dadosCep.cidade;
                currentSheet.Cells[currentRow, ++currentColumn] = dadosCep.uf;
                currentSheet.Cells[currentRow, ++currentColumn] = dadosCep.unidadesPostagem;

                currentRow++;
            }
        }

        //@Todo: this should have an specific data structure
        public void toCurrentPlan(List<dynamic> dadosCnpj)
        {
            int currentRow = currentSelection.Cells[1, 1].row; //use this instead of currentCell to avoid misspositioning
            foreach (dynamic dados in dadosCnpj)
            {
                int currentColumn = currentSelection.Cells[1, 1].Column; //use this instead of currentCell to avoid misspositioning

                currentSheet.Cells[currentRow, ++currentColumn] = dados.matriz_filial;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.razao_social;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.nome_fantasia;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.situacao;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.data_situacao;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.motivo_situacao;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.nm_cidade_exterior;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.cod_pais;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.nome_pais;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.cod_nat_juridica;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.data_inicio_ativ;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.cnae_fiscal;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.tipo_logradouro;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.logradouro;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.numero;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.complemento;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.bairro;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.cep;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.uf;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.municipio;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.ddd_1;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.telefone_1;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.ddd_2;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.telefone_2;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.ddd_fax;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.num_fax;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.email;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.qualif_resp;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.porte;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.opc_simples;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.data_opc_simples;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.data_exc_simples;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.opc_mei;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.sit_especial;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.data_sit_especial;
                currentSheet.Cells[currentRow, ++currentColumn] = dados.capital_social;
                currentRow++;
            }

        }

        private void btnConsultaCNPJSClip_Click(object sender, RibbonControlEventArgs e)
        {
            BuscaCnpjFromWs(true);
        }

        private void btnConsultaCEPClip_Click(object sender, RibbonControlEventArgs e)
        {
            BuscaCEPFromCorreioWs(true);
        }
    }
}
