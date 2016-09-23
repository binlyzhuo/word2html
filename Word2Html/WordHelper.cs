using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Aspose.Words;

namespace Word2Html
{
    public static class WordHelper
    {
        public static void HtmlToWord(string htmlContent)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertHtml(htmlContent);
            doc.Save("d:\\111.doc");
        }
    }
}