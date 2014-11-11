using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace LinkValidator
{
    public class Program
    {
        public static void Main(String[] args)
        {
            var docPath = @"D:\Documents\OpenXMLTest\test2.docx";

            using (WordprocessingDocument doc = WordprocessingDocument.Open(docPath, false))
            {
                ProcessHyperlinks(doc);
            }

            Console.Read();
        }

        private static void ProcessHyperlinks(WordprocessingDocument doc)
        {
            var mainDocument = doc.MainDocumentPart.Document;
            var links = mainDocument.MainDocumentPart.Document.Body.Descendants<Hyperlink>();

            Console.WriteLine(String.Format("Found {0} hyperlinks.{1}", links.Count(), Environment.NewLine));

            foreach (var hyperlink in links)
            {
                var hyperlinkText = new StringBuilder();

                foreach (Text text in hyperlink.Descendants<Text>())
                    hyperlinkText.Append(text.InnerText);

                var hyperlinkRelationshipId = hyperlink.Id.Value;
                var hyperlinkRelationship = mainDocument.MainDocumentPart.HyperlinkRelationships.First(r => r.Id == hyperlinkRelationshipId);

                var status = String.Empty;
                var uri = String.Empty;

                if (hyperlinkRelationship.Uri.Scheme == "http" || hyperlinkRelationship.Uri.Scheme == "https")
                {
                    uri = hyperlinkRelationship.Uri.AbsoluteUri;
                    var request = WebRequest.Create(uri);

                    try
                    {
                        var response = (HttpWebResponse)request.GetResponse();
                        status = response.StatusCode.ToString();
                    }
                    catch (WebException ex)
                    {
                        status = ex.Status.ToString();
                    }
                }
                else if (hyperlinkRelationship.Uri.Scheme == "file")
                {
                    if (hyperlinkRelationship.Uri.IsUnc)
                        uri = hyperlinkRelationship.Uri.LocalPath;
                    else
                        uri = hyperlinkRelationship.Uri.AbsolutePath;

                    if (File.Exists(uri))
                        status = "File Exists";
                    else
                        status = "File Does Not Exist";
                }
                

                Console.WriteLine(String.Format("Link Text: {0}", hyperlinkText));
                Console.WriteLine(String.Format("Link address: {0}", uri));
                Console.WriteLine(String.Format("Status: {0}", status));
                Console.WriteLine(Environment.NewLine);
            }
        }
    }
}
