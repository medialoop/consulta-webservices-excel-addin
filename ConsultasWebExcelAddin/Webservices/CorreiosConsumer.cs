using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using ConsultasWebExcelAddin.wsCorreios;

namespace ConsultasWebExcelAddin.WebService
{


    public class CorreiosConsumer
    {
        public static enderecoERP getFullAddressFromCorreios(string cep)
        {

            Regex removeCepChars = new Regex("[^0-9]");
            cep = removeCepChars.Replace(cep, "");
            
            if(cep.Length < 7 || cep.Length > 8)
            {
                throw new System.Exception("CEP deve ter 7 ou 8 caracteres");
            }

            AtendeClienteClient ws = new AtendeClienteClient();
            enderecoERP dadosCep = ws.consultaCEP(cep);

            return dadosCep;
        }
    }
}
