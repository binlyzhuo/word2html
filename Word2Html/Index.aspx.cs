using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using HtmlAgilityPack;
using Aspose.Words;


namespace Word2Html
{
    public partial class Index : System.Web.UI.Page
    {
        private string saveDoc = "D:\\mydocs";
        protected void Page_Load(object sender, EventArgs e)
        {
            string wordFile = "";

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            string fileName = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            FileUpload1.SaveAs(string.Format(@"{0}\\{1}.doc",saveDoc,fileName));

            Aspose.Words.Document d = new Aspose.Words.Document(string.Format(@"{0}\\{1}.doc", saveDoc, fileName));
            d.Save(string.Format(@"{0}\\{1}.html", saveDoc, fileName), SaveFormat.Html);


            HyperLink1.NavigateUrl = string.Format("/HtmlPage.aspx?f={0}",fileName);

            HtmlDocument document = new HtmlDocument();
            document.Load(string.Format(@"{0}\\{1}.html", saveDoc, fileName),Encoding.UTF8);

            var docNode= document.DocumentNode.SelectNodes("/html/body/div/p");


            foreach (var dc in docNode)
            {
                Response.Write(dc.InnerText);
                Response.Write("<br/>");
            }

        }
    }
}