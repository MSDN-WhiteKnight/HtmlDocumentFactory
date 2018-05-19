using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Windows.Forms;
using HtmlLib;
/* Project: HtmlDocumentFactory
 * Author: MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight)
 * License: BSD 3-Clause
 */

namespace Html
{
    /********  HtmlDocumentFactory usage sample  ********/
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            HtmlDocument htmldoc;
            htmldoc = HtmlDocumentFactory.CreateHtmlDocument();

            //заполним содержимое документа
            htmldoc.Title = "Hello";

            HtmlElement el = htmldoc.CreateElement("h1");
            el.InnerText = "Hello, world!";
            htmldoc.Body.AppendChild(el);

            el = htmldoc.CreateElement("div");
            el.InnerText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.";
            el.Style = "color: red";
            htmldoc.Body.AppendChild(el);

            //получаем все содержимое документа в виде html
            textBox1.Text = HtmlDocumentFactory.HtmlDocumentToString(htmldoc);

            //освобождаем неуправляемые ресурсы, связанные с HtmlDocument
            HtmlDocumentFactory.ReleaseHtmlDocument(htmldoc);
            
        }
    }
}
