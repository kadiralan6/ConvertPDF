using Newtonsoft.Json;
using SelectPdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ConvertPDF.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult UrasPdf()
        {
            return View();
        }

        [HttpPost]
        public string UploadExcel(List<ExcelUpload> uploadList)
        {

            List<ChangePDFItems> ruleList = new List<ChangePDFItems>();
            foreach (var uploadItem in uploadList)
            {

                ChangePDFItems items = new ChangePDFItems();
                items.Mahiyet = uploadItem.text;
                items.IrsaliyeTarihi = uploadItem.DateTime;
                items.Address = uploadItem.Address;
                items.SerialNo = uploadItem.SerialNo;
                items.EVRAKNO = uploadItem.evrak;
                items.TCNumber = uploadItem.TCNumber;
                items.CustomerName = uploadItem.CustomerName;
                items.Quantity = uploadItem.Quantity;
                items.PriceList = uploadItem.PriceList;
                items.TotalPrice = uploadItem.TotalPrice;

                ruleList.Add(items);

            }
            return JsonConvert.SerializeObject(ruleList);

        }
        [HttpPost]
        public string SavePdf(List<ChangePDFItems> pdfItems, int type)
        {
            byte[] pdfByte = null;
            string file = string.Empty;
            string contentHtml = "";
            string path = "";
            string pdfPath = null;
            string marge = null;

            var btylist = new List<byte[]>();
            int i = 1,counts=1;
            foreach (var item in pdfItems)
            {
                if(type==1)
                file = System.Web.Hosting.HostingEnvironment.MapPath("~/files/template/irsaliye.html");
                else
                    file = System.Web.Hosting.HostingEnvironment.MapPath("~/files/template/irsaliye2.html");
                path = "Files/" + Guid.NewGuid() + ".pdf";
                contentHtml = System.IO.File.ReadAllText(file);
                contentHtml = contentHtml.Replace("{Works}", item.Mahiyet)
                                          .Replace("{Quantity}", item.Quantity.ToString("N2"))
                                          .Replace("{CustomerName}", item.CustomerName)
                                          .Replace("{TCNO}", item.TCNumber)
                                          .Replace("{Date}", item.DateTimeStr)
                                          .Replace("{Address}", item.Address)
                                          .Replace("{EvrakNo}", item.EVRAKNO)
                                          .Replace("{PriceList}", item.PriceList.ToString("N2") + "TL")
                                          .Replace("{TotalPrice}", item.TotalPrice.ToString("N2") + "TL")
                                          .Replace("{marginTop}", i==2 ? "marginTop":"")
                                          ;
                
                HtmlToPdf converter = new HtmlToPdf();
                converter.Options.PdfPageSize = PdfPageSize.A4;
                converter.Options.MarginLeft = 40;
                converter.Options.MarginRight = 40;
                converter.Options.MarginTop = 10;

                converter.Options.MarginBottom = 10;
                marge += contentHtml;
                if (i == 2 || counts==pdfItems.Count)
                {

                    SelectPdf.PdfDocument doc = converter.ConvertHtmlString(marge, string.Empty);

                    // save pdf document
                    byte[] pdf = doc.Save();

                    // close pdf document
                    doc.Close();

                    btylist.Add(pdf);
                    i = 0; marge = "";
                }
                i++;
                counts++;
                // create a new pdf document converting an url




            }
            byte[] combinedPdf = CombinePDFs(btylist);
            //https://localhost:44300/
            Upload(Server.MapPath("~/" + path), combinedPdf);

            pdfPath = "http://arengeridonusum.com/" + path;
            //return JsonConvert.SerializeObject(pdfPath);
            return pdfPath;

        }
        public static bool Upload(string path, byte[] data)
        {
            try
            {
                System.IO.FileInfo file = new System.IO.FileInfo(path);
                file.Directory.Create();

                System.IO.File.WriteAllBytes(path, data);
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }
        public static byte[] CombinePDFs(List<byte[]> srcPDFs)
        {
            using (var ms = new MemoryStream())
            {
                using (var resultPDF = new PdfSharp.Pdf.PdfDocument(ms))
                {
                    foreach (var pdf in srcPDFs)
                    {
                        using (var src = new MemoryStream(pdf))
                        {
                            using (var srcPDF = PdfReader.Open(src, PdfDocumentOpenMode.Import))
                            {
                                for (var i = 0; i < srcPDF.PageCount; i++)
                                {
                                    int a = i % 2;
                                    if (a == 0)
                                        resultPDF.AddPage(srcPDF.Pages[i]);
                                }
                            }
                        }
                    }
                    resultPDF.Save(ms);
                    return ms.ToArray();
                }
            }
        }
        public class ExcelUpload
        {
            public DateTime DateTime { get; set; }
            public string CustomerName { get; set; }
            public string TCNumber { get; set; }
            public string Address { get; set; }
            public string SerialNo { get; set; }
            public string evrak { get; set; }
            public string text { get; set; }
          
      
            public double Quantity { get; set; }
            public double PriceList { get; set; }
            public double TotalPrice { get; set; }

        }
        public class ChangePDFItems
        {
        
            public DateTime IrsaliyeTarihi { get; set; }

            public string DateTimeStr { get { return IrsaliyeTarihi.ToString("dd/MM/yyyy"); } }
            public string CustomerName { get; set; }
            public string TCNumber { get; set; }
            public string Address { get; set; }
            public string Mahiyet { get; set; }
            public string SerialNo { get; set; }
            public string EVRAKNO { get; set; }
            public double Quantity { get; set; }
            public double PriceList { get; set; }
            public double TotalPrice { get; set; }

        }

    }
}