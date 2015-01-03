<%@ WebHandler Language="C#" Class="Handler" %>

using System;
using System.IO;
using System.IO.Compression;
using System.Web;
using System.Web.Script.Serialization;

/*
 * This handler will accept the document data from the client, 
 * extract the main data from the docx (which is a zip archive) and return it to the client.
 */
public class Handler : IHttpHandler
{
    public void ProcessRequest(HttpContext context)
    {
        if (context.Request.InputStream.Length > 0)
        {
            string webRequestData = "";
            using (StreamReader sr = new StreamReader(context.Request.InputStream))
            {
                webRequestData = sr.ReadToEnd();
            }
            byte[] base64String = Convert.FromBase64String(webRequestData);
            Stream docxStream = new MemoryStream(base64String);
            try
            {
                ZipArchive archive = new ZipArchive(docxStream);

                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    // The text content of a docx is stored in word/document.xml. So we seek the data from there.
                    // You may do further operation with OpenXmlApi.
                    if (entry.FullName == "word/document.xml")
                    {
                        Stream documentXmlData = entry.Open();
                        StreamReader reader = new StreamReader(documentXmlData);
                        string text = reader.ReadToEnd();
                        context.Response.Write(text);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                context.Response.Write(ex.Message);
            }
        }
        context.Response.ContentType = "text/plain";
    }

    public bool IsReusable
    {
        get
        {
            return false;
        }
    }
}