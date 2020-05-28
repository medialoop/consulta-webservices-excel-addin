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

namespace ConsultasWebExcelAddin
{
    public partial class RibbonMain
    {
        private void RibbonMain_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnConsultaCNPJSCelulas_Click(object sender, RibbonControlEventArgs e)
        {
            BuscaCnpjFromWs();
        }

        private void BuscaCnpjFromWs()
        {
            try
            {
                List<dynamic> DadosCnpj;

                Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
                Range currentCell = Globals.ThisAddIn.getCurrentCell();
                Range currentSeleciton = Globals.ThisAddIn.getSelectedRange();
                List<string> cnpjs = new List<string>();

                foreach (Range cellCnpj in Globals.ThisAddIn.getSelectedRange())
                {
                    if (cellCnpj.Value2 != null)
                    {
                       cnpjs.Add(cellCnpj.Value2.ToString());
                    }
                }

                DialogResult dialogResult = MessageBox.Show("Está operação não é reversível", 
                                                            "Está operação preenche as 36 colunas imediatamente depois dos CNPJs selecionados, deseja continuar?", 
                                                            MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.No)
                {
                    return;
                }

                DadosCnpj = TiCnpjConsumer.getFullDataByCnpj(cnpjs);

                int cRow = currentCell.Column;

                foreach (dynamic dados in DadosCnpj)
                {
                    int cColumn = currentCell.Column;

                    currentSheet.Cells[cRow, ++cColumn] = dados.matriz_filial;
                    currentSheet.Cells[cRow, ++cColumn] = dados.razao_social;
                    currentSheet.Cells[cRow, ++cColumn] = dados.nome_fantasia;
                    currentSheet.Cells[cRow, ++cColumn] = dados.situacao;
                    currentSheet.Cells[cRow, ++cColumn] = dados.data_situacao;
                    currentSheet.Cells[cRow, ++cColumn] = dados.motivo_situacao;
                    currentSheet.Cells[cRow, ++cColumn] = dados.nm_cidade_exterior;
                    currentSheet.Cells[cRow, ++cColumn] = dados.cod_pais;
                    currentSheet.Cells[cRow, ++cColumn] = dados.nome_pais;
                    currentSheet.Cells[cRow, ++cColumn] = dados.cod_nat_juridica;
                    currentSheet.Cells[cRow, ++cColumn] = dados.data_inicio_ativ;
                    currentSheet.Cells[cRow, ++cColumn] = dados.cnae_fiscal;
                    currentSheet.Cells[cRow, ++cColumn] = dados.tipo_logradouro;
                    currentSheet.Cells[cRow, ++cColumn] = dados.logradouro;
                    currentSheet.Cells[cRow, ++cColumn] = dados.numero;
                    currentSheet.Cells[cRow, ++cColumn] = dados.complemento;
                    currentSheet.Cells[cRow, ++cColumn] = dados.bairro;
                    currentSheet.Cells[cRow, ++cColumn] = dados.cep;
                    currentSheet.Cells[cRow, ++cColumn] = dados.uf;
                    currentSheet.Cells[cRow, ++cColumn] = dados.municipio;
                    currentSheet.Cells[cRow, ++cColumn] = dados.ddd_1;
                    currentSheet.Cells[cRow, ++cColumn] = dados.telefone_1;
                    currentSheet.Cells[cRow, ++cColumn] = dados.ddd_2;
                    currentSheet.Cells[cRow, ++cColumn] = dados.telefone_2;
                    currentSheet.Cells[cRow, ++cColumn] = dados.ddd_fax;
                    currentSheet.Cells[cRow, ++cColumn] = dados.num_fax;
                    currentSheet.Cells[cRow, ++cColumn] = dados.email;
                    currentSheet.Cells[cRow, ++cColumn] = dados.qualif_resp;
                    currentSheet.Cells[cRow, ++cColumn] = dados.porte;
                    currentSheet.Cells[cRow, ++cColumn] = dados.opc_simples;
                    currentSheet.Cells[cRow, ++cColumn] = dados.data_opc_simples;
                    currentSheet.Cells[cRow, ++cColumn] = dados.data_exc_simples;
                    currentSheet.Cells[cRow, ++cColumn] = dados.opc_mei;
                    currentSheet.Cells[cRow, ++cColumn] = dados.sit_especial;
                    currentSheet.Cells[cRow, ++cColumn] = dados.data_sit_especial;
                    currentSheet.Cells[cRow, ++cColumn] = dados.capital_social;
                    cRow++;
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
                MessageBox.Show(ex.Message, "Ocorreu um erro ao buscar informação do CNPJ", MessageBoxButtons.OK);
            }
        }


        private void btnBuscarCEPCelulas_Click(object sender, RibbonControlEventArgs e)
        {
            BuscaCEPFromCorreioWs();
        }

        private void BuscaCEPFromCorreioWs()
        {
            try
            {
                Excel.Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
                Excel.Range currentCell = Globals.ThisAddIn.getCurrentCell();
                string cepValue = Convert.ToString(currentCell.Value2);

                if (cepValue is null)
                {
                    throw new System.Exception("Selecione um campo contendo um CEP válido");
                }

                enderecoERP dadosCep = CorreiosConsumer.getFullAddressFromCorreios(cepValue);

                int currentCollumn = currentCell.Column;
                currentSheet.Cells[currentCell.Row, ++currentCollumn] = dadosCep.end;
                currentSheet.Cells[currentCell.Row, ++currentCollumn] = dadosCep.bairro;
                currentSheet.Cells[currentCell.Row, ++currentCollumn] = dadosCep.cidade;
                currentSheet.Cells[currentCell.Row, ++currentCollumn] = dadosCep.uf;
                currentSheet.Cells[currentCell.Row, ++currentCollumn] = dadosCep.unidadesPostagem;


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
        }
    }
}
