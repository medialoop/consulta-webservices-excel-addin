using Newtonsoft.Json;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;

namespace ConsultasWebExcelAddin
{

    public class TiCnpjConsumer
    {
        private static string wsUrlBase = "http://localhost:8080/empresas";
        private static string wsUrlGetOneCnpj = wsUrlBase + "/cnpj/";
        private static string wsUrlSearchRazaoCnpj = wsUrlBase + "/razao/";


        public static dynamic getFullDataByCnpj(string Cnpj)
        {
            Regex removeCnpjChars = new Regex("[^0-9]");
            Cnpj = removeCnpjChars.Replace(Cnpj, "");

            if (Cnpj.Length != 14)
            {
                throw new System.Exception("Cnpj em formato errado, mínimo de 14 caracteres");
            }

            WebRequest request = WebRequest.Create(wsUrlGetOneCnpj + Cnpj);
            request.Method = "GET";
            request.ContentLength = 0;
            request.ContentType = "application/json";

            string jsonResponse;
            using (WebResponse response = request.GetResponse())
            {
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    jsonResponse = reader.ReadToEnd();
                }
            }

            dynamic rootList;
            rootList = JsonConvert.DeserializeObject(jsonResponse);

            return rootList;
        }

    }
}