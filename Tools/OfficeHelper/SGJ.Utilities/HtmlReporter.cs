using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;
using System.IO;

namespace SGJ.Utilities
{

    public class Parameter
    {
        public string Name { get; set; }
        public bool IsSuccess { get; set; }
        public string Remark { get; set; }
    }
    public class HtmlReporter
    {

        public static void WriteBR(HtmlTextWriter writer)
        {
            writer.WriteBeginTag("br");
            writer.Write(HtmlTextWriter.SelfClosingTagEnd);
            //or
            writer.Write("<br />");
        }


        public static string Generate(List<Parameter> lcolPara)
        {
            StringWriter stringWriter = new StringWriter();
            using (HtmlTextWriter writer = new HtmlTextWriter(stringWriter))
            {
                writer.RenderBeginTag(HtmlTextWriterTag.Html);
                writer.RenderBeginTag(HtmlTextWriterTag.Head);
                writer.RenderBeginTag(HtmlTextWriterTag.Title);
                writer.Write("Export Result");
                writer.RenderEndTag(); // End title
                writer.AddAttribute("http-equiv", "Content-Type");
                writer.AddAttribute("content", "text/html; charset=utf-8");
                writer.RenderBeginTag(HtmlTextWriterTag.Meta);
                writer.RenderEndTag(); //End meta
                writer.RenderEndTag(); //End head
                writer.RenderBeginTag(HtmlTextWriterTag.Body);// Begin Body
                writer.AddAttribute("face", "Arial");
                writer.RenderBeginTag(HtmlTextWriterTag.Font); // Begin Font
                writer.Write("Export Result");
                writer.RenderEndTag(); // End H2
                writer.RenderBeginTag(HtmlTextWriterTag.Hr); // Begin Hr
                writer.RenderEndTag(); // End Hr
                writer.RenderBeginTag(HtmlTextWriterTag.P); // Begin P
                writer.AddAttribute("cellpadding", "3");
                writer.RenderBeginTag(HtmlTextWriterTag.Table); // Begin Table
                foreach (Parameter para in lcolPara)
                {
                    AddTableDetail(writer, para);
                }
                writer.RenderEndTag(); // End Table
                writer.RenderEndTag(); // End P
                writer.RenderEndTag(); // End Font
                writer.RenderEndTag(); // End Body

            }
            // Return the result.
            return stringWriter.ToString();
        }

        public static void AddTableDetail(HtmlTextWriter writer, Parameter para)
        {
            writer.AddAttribute("bgcolor", "#808080");
            writer.RenderBeginTag(HtmlTextWriterTag.Tr); // begin tr
            //td 1
            writer.AddAttribute("style", "font-weight: bold; color: white;");
            writer.RenderBeginTag(HtmlTextWriterTag.Td); // Begin Td
            writer.Write(para.Name);
            writer.RenderEndTag(); // End Td
            //td 2
            writer.AddAttribute("style", "font-weight: bold; color: white;");
            writer.AddAttribute("align", "center");
            writer.RenderBeginTag(HtmlTextWriterTag.Td); // Begin Td
            writer.Write(para.IsSuccess);
            writer.RenderEndTag(); // End Td
            //td 3
            writer.AddAttribute("style", "font-weight: normal; color: white;");
            writer.RenderBeginTag(HtmlTextWriterTag.Td); // Begin Td
            writer.Write(para.Remark);
            writer.RenderEndTag(); // End Td
            writer.RenderEndTag(); // end tr
        }
    }
}
