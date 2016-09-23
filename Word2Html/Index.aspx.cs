using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using HtmlAgilityPack;
using Aspose.Words;
using Microsoft.SqlServer.Server;


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
            string wordFileName = fileName + ".doc";
            string htmlFileName = fileName+".html";

            string tempFolder = fileName+new Random().Next(100);
            if (!Directory.Exists(string.Format("{0}\\{1}", saveDoc, tempFolder)))
            {
                Directory.CreateDirectory(string.Format("{0}\\{1}", saveDoc, tempFolder));
            }

            FileUpload1.SaveAs(string.Format(@"{0}\\{1}\\{2}",saveDoc,tempFolder,wordFileName));

            Aspose.Words.Document d = new Aspose.Words.Document(string.Format(@"{0}\\{1}\\{2}", saveDoc,tempFolder, wordFileName));
            d.Save(string.Format(@"{0}\\{1}\\{2}", saveDoc,tempFolder, htmlFileName), SaveFormat.Html);


            HyperLink1.NavigateUrl = string.Format("/HtmlPage.aspx?f={0}",wordFileName);

            HtmlDocument document = new HtmlDocument();
            document.Load(string.Format(@"{0}\\{1}\\{2}", saveDoc,tempFolder, htmlFileName),Encoding.UTF8);

            var docNode= document.DocumentNode.SelectNodes("/html/body/div/p");


            foreach (var dc in docNode)
            {
                Response.Write(dc.InnerText);
                var imgNodes = dc.SelectNodes("img");
                if (imgNodes!=null&&imgNodes.Count > 0)
                {
                    foreach (var im in imgNodes)
                    {
                        var imgUrl = im.Attributes["src"];
                        string imgPath = string.Format("{0}\\{1}\\{2}", saveDoc, tempFolder, imgUrl.Value);
                        Response.Write("<br/>");
                        string base64 = ImageHelper.ImageToBase64(imgPath);
                        Response.Write(base64);
                        Response.Write("<br/>");
                        Response.Write("<img src='data:image/png;base64," + base64 + "'/>");

                    }
                }
                Response.Write("<br/>");
            }

            
        }
    }
}