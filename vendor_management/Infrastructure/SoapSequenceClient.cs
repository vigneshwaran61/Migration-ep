using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Extensions.Configuration;

namespace SYS_VENDOR_MGMT.Infrastructure
{
    /// <summary>
    /// Replaces the legacy ASMX Web Reference proxies:
    ///   Vendor_SEQ.VENDOR_SEQ  →  GetSvmPaymentGeneratedDtlSeqAsync
    ///   Vendor_SEQ.VENDOR_SEQ  →  GetSvmVendorNetpayDtlSeqAsync
    ///
    /// Calls are made via HttpClient posting a minimal SOAP 1.1 envelope.
    /// TLS 1.2/1.3 is enforced globally in Program.cs via ServicePointManager.
    /// </summary>
    public class SoapSequenceClient
    {
        private readonly IHttpClientFactory _factory;
        private readonly string _vendorSeqUrl;

        public SoapSequenceClient(IHttpClientFactory factory, IConfiguration config)
        {
            _factory       = factory;
            _vendorSeqUrl  = config["AppSettings:VendorSeqServiceUrl"]
                             ?? throw new InvalidOperationException("VendorSeqServiceUrl not configured.");
        }

        public Task<string> GetSvmPaymentGeneratedDtlSeqAsync() =>
            CallSoapMethodAsync("GET_SVM_PAYMENT_GENRATED_DTL_SEQ");

        public Task<string> GetSvmVendorNetpayDtlSeqAsync() =>
            CallSoapMethodAsync("GET_SVM_VENDORNETPAY_DTL_SEQ");

        private async Task<string> CallSoapMethodAsync(string methodName)
        {
            string soapEnvelope =
                $"""
                <?xml version="1.0" encoding="utf-8"?>
                <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                  <soap:Body>
                    <{methodName} xmlns="http://tempuri.org/" />
                  </soap:Body>
                </soap:Envelope>
                """;

            var client  = _factory.CreateClient();
            var content = new StringContent(soapEnvelope, Encoding.UTF8, "text/xml");
            content.Headers.Add("SOAPAction", $"\"http://tempuri.org/{methodName}\"");

            var response = await client.PostAsync(_vendorSeqUrl, content);
            response.EnsureSuccessStatusCode();

            string xml = await response.Content.ReadAsStringAsync();
            var doc     = new XmlDocument();
            doc.LoadXml(xml);
            // Navigate to the first non-whitespace text inside the Body
            XmlNamespaceManager ns = new(doc.NameTable);
            ns.AddNamespace("soap", "http://schemas.xmlsoap.org/soap/envelope/");
            XmlNode? resultNode = doc.SelectSingleNode($"//soap:Body//{methodName}Result", ns)
                                  ?? doc.SelectSingleNode($"//*[local-name()='{methodName}Result']");
            return resultNode?.InnerText ?? "";
        }
    }
}
