# HtmlDocumentFactory

Windows Forms framework has a nice `HtmlDocument` class for HTML manupulation. However, it can only work with documents from `WebBrowser` control, and there's no public constructor to create new `HtmlDocument` instance from scratch. This library attempts to fix this problem. It consists of a single static class `HtmlLib.HtmlDocumentFactory` that provides helper methods for creating and destroying HTML documents and converting them to strings.

**Usage**

Add reference to *HtmlDocumentFactory.dll* and import namespace with `using HtmlLib;` clause.

Create a new instance of HtmlDocument:

    HtmlDocument htmldoc = HtmlDocumentFactory.CreateHtmlDocument();

Modify document's content like you wish:

    HtmlElement el = htmldoc.CreateElement("div");
    el.InnerText = "Hello, world!";
    el.Style = "color: red";
    htmldoc.Body.AppendChild(el);

Copy resulting HTML to string object:

    textBox1.Text = HtmlDocumentFactory.HtmlDocumentToString(htmldoc);
    
Free unmanaged resources when you no longer need the document:

    HtmlDocumentFactory.ReleaseHtmlDocument(htmldoc);
