using Microsoft.Office.Tools.Ribbon;
using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TiBuscaCnpjWS
{
    public partial class RibbonTiCnpjWS
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void BtnCallWsIntoPlan_Click(object sender, RibbonControlEventArgs e)
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

                Regex removeCnpjChars = new Regex("[^0-9]");
                cnpjValue = removeCnpjChars.Replace(cnpjValue, "");

                if (cnpjValue.Length != 14)
                {
                    throw new Exception("Cnpj em formato errado, mínimo de 14 caracteres");
                }

                DadosCnpj = WebServiceConsumer.getFullDataByCnpj(cnpjValue);

                currentSheet.Cells[currentCell.Row, currentCell.Column + 1] = DadosCnpj.razao_social;
                currentSheet.Cells[currentCell.Row, currentCell.Column + 2] = DadosCnpj.nome_fantasia;

                //\DialogResult result = MessageBox.Show(DadosCnpj.ToString(), "erro", MessageBoxButtons.OK);

            }
            catch (System.Net.WebException ex)
            {
                MessageBox.Show("Verifique sua conexão com a internet" + ex.Message,
                                "Consulta CNPJ - Erro de Rede!",
                                MessageBoxButtons.OK);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ocorreu um erro ao buscar informação do CNPJ", MessageBoxButtons.OK);
            }
        }

    }
}
