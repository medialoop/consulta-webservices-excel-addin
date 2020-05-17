using Newtonsoft.Json;
using System.IO;
using System.Net;

namespace TiBuscaCnpjWS
{

    public class WebServiceConsumer
    {
        static string wsUrlBase = "http://localhost:8080/empresas";
        static string wsUrlGetOneCnpj = wsUrlBase + "/cnpj/";
        static string wsUrlSearchRazaoCnpj = wsUrlBase + "/razao/";

        public static dynamic getFullDataByCnpj(string Cnpj)
        {
            WebRequest request = WebRequest.Create(WebServiceConsumer.wsUrlGetOneCnpj + Cnpj);
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