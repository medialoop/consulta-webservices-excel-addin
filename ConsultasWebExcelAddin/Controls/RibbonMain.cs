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

namespace ConsultasWebExcelAddin
{
    public partial class RibbonMain
    {
        private void RibbonMain_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnConsultaCNPJSCelulas_Click(object sender, RibbonControlEventArgs e)
        {
            this.BuscaCnpjFromWs();
        }

        private void BuscaCnpjFromWs()
        {
            try
            {
                dynamic DadosCnpj;

                Excel.Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
                Excel.Range currentCell = Globals.ThisAddIn.getCurrentCell();
                string cnpjValue = Convert.ToString(currentCell.Value2);
 

                if (cnpjValue is null)
                {
                    throw new System.Exception("Selecione um campo contendo um CNPJ válido");
                }

                DadosCnpj = TiCnpjConsumer.getFullDataByCnpj(cnpjValue);

                currentSheet.Cells[currentCell.Row, currentCell.Column + 1] = DadosCnpj.razao_social;
                currentSheet.Cells[currentCell.Row, currentCell.Column + 2] = DadosCnpj.nome_fantasia;

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
                MessageBox.Show(ex.Message, "Ocorreu um erro ao buscar informação do CEP: "+ex.Message, MessageBoxButtons.OK);
            }

        }
    }
}
