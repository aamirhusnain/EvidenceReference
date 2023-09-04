using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Mvc;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace EvidenceReferenceWeb.Controllers
{
    public class PDFTestingController : Controller
    {
        public class ObjectClass
        {
            public string Exhibitid { get; set; }
            public string Affiant { get; set; }
            public string Commissioner { get; set; }
            public string Date { get; set; }
            public string Link { get; set; }
        }

        public class PdfConversionData
        {
            public string Base64Data { get; set; }
            public List<ObjectClass> ObjectsArray { get; set; }
        }

        [HttpPost]
        public ActionResult ConvertToPdf([System.Web.Http.FromBody] PdfConversionData pdfData)
        {
            try
            {
                // Decode the base64 data
                byte[] decodedBytes = Convert.FromBase64String(pdfData.Base64Data);

                // Create a MemoryStream from the decoded data
                using (MemoryStream stream = new MemoryStream(decodedBytes))
                {
                    // Create a new document
                    Document document = new Document();

                 
                    Response.ContentType = "application/pdf";

                    // Provide a filename for the downloaded PDF
                    Response.AddHeader("Content-Disposition", "attachment; filename=converted-document.pdf");

                    
                    PdfWriter writer = PdfWriter.GetInstance(document, Response.OutputStream);

                    // Open the document for writing
                    document.Open();

                    // Create a PDF reader for the content
                    PdfReader reader = new PdfReader(stream);

                    
                    for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; pageNumber++)
                    {
                        document.NewPage();
                        PdfImportedPage page = writer.GetImportedPage(reader, pageNumber);
                        writer.DirectContent.AddTemplate(page, 0, 0);
                    }

                    // Add each object from the array to a new page

                    foreach (var exhibit in pdfData.ObjectsArray)
                    {
                        document.NewPage();
                        Paragraph paragraph = new Paragraph();
                        paragraph.Alignment = Element.ALIGN_CENTER; // Center align the text
                        paragraph.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12); // Set the font to bold

                        // Add the text for the current object to the paragraph
                        paragraph.Add(new Chunk($"ThIS IS {exhibit.Exhibitid} REFERRED\n", FontFactory.GetFont(FontFactory.HELVETICA, 12)));
                        paragraph.Add(new Chunk($"TO IN THE AFFIDAVIT OF {exhibit.Affiant} \n", FontFactory.GetFont(FontFactory.HELVETICA, 12)));
                        paragraph.Add(new Chunk($"SWORN THIS {exhibit.Date}\n", FontFactory.GetFont(FontFactory.HELVETICA, 12)));
                        paragraph.Add(new Chunk($"\n", FontFactory.GetFont(FontFactory.HELVETICA, 12)));
                        paragraph.Add(new Chunk($"\n", FontFactory.GetFont(FontFactory.HELVETICA, 12)));
                        paragraph.Add(new Chunk($"A Commissioner for Taking Affidavits, etc.\n", FontFactory.GetFont(FontFactory.HELVETICA, 12)));
                        paragraph.Add(new Chunk($"{exhibit.Commissioner}\n", FontFactory.GetFont(FontFactory.HELVETICA, 12)));
                      //paragraph.Add(new Chunk($"Data: {exhibit.Link}\n", FontFactory.GetFont(FontFactory.HELVETICA, 12)));

                        document.Add(paragraph);
                    }

                    // Close the document
                    document.Close();
                }

                Response.End(); 

                return new EmptyResult(); 
            }
            catch (Exception ex)
            {
                // Handle any errors or exceptions here
                return Content("Error: " + ex.Message);
            }
        }
    }
}

//public class MyDataObject
//{
//    public string Exhibitid { get; set; }
//    public string data { get; set; }
//}

//public class MyDataModel
//{
//    public string Lead { get; set; }
//    public string Exhibit { get; set; }
//    public string Description { get; set; }
//    public string Link { get; set; }
//}
