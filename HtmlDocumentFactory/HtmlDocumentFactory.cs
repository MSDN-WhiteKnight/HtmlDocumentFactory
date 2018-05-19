using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
/* Project: HtmlDocumentFactory
 * Author: MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight)
 * License: BSD 3-Clause
 */

namespace HtmlLib
{
    [ComImport(), Guid("25336920-03F9-11CF-8FD0-00AA00686F13")]
    class HTMLDocument
    {
    }
        
    [ComImport, Guid("626FC520-A41E-11CF-A731-00A0C9082637"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    interface IHTMLDocument
    {             
        [return: MarshalAs(UnmanagedType.IDispatch)]
        object GetScript();
    }

    [ComImport, Guid("332C4425-26CB-11D0-B483-00C04FD90119"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    interface IHTMLDocument2 : IHTMLDocument
    {       
        [return: MarshalAs(UnmanagedType.Interface)]
        new object GetScript();        
        dynamic GetAll();
        [return: MarshalAs(UnmanagedType.Interface)]        
        object GetBody();
        [return: MarshalAs(UnmanagedType.Interface)]        
        object GetActiveElement();
        dynamic GetImages();
        dynamic GetApplets();
        dynamic GetLinks();
        dynamic GetForms();
        dynamic GetAnchors();
        void SetTitle(string p);
        string GetTitle();
        dynamic GetScripts();
        void SetDesignMode(string p);
        string GetDesignMode();
        [return: MarshalAs(UnmanagedType.Interface)]
        object GetSelection();
        string GetReadyState();
        [return: MarshalAs(UnmanagedType.Interface)]
        object GetFrames();
        dynamic GetEmbeds();
        dynamic GetPlugins();
        void SetAlinkColor(object c);
        object GetAlinkColor();
        void SetBgColor(object c);
        object GetBgColor();
        void SetFgColor(object c);
        object GetFgColor();
        void SetLinkColor(object c);
        object GetLinkColor();
        void SetVlinkColor(object c);
        object GetVlinkColor();
        string GetReferrer();
        dynamic GetLocation();
        string GetLastModified();
        void SetUrl(string p);
        string GetUrl();
        void SetDomain(string p);
        string GetDomain();
        void SetCookie(string p);
        string GetCookie();
        void SetExpando(bool p);
        bool GetExpando();
        void SetCharset(string p);
        string GetCharset();
        void SetDefaultCharset(string p);
        string GetDefaultCharset();
        string GetMimeType();
        string GetFileSize();
        string GetFileCreatedDate();
        string GetFileModifiedDate();
        string GetFileUpdatedDate();
        string GetSecurity();
        string GetProtocol();
        string GetNameProp();
        int Write([In, MarshalAs(UnmanagedType.SafeArray)] object[] psarray);
        int WriteLine([In, MarshalAs(UnmanagedType.SafeArray)] object[] psarray);
        [return: MarshalAs(UnmanagedType.Interface)]
        object Open(string mimeExtension, object name, object features, object replace);
        void Close();
        void Clear();
        bool QueryCommandSupported(string cmdID);
        bool QueryCommandEnabled(string cmdID);
        bool QueryCommandState(string cmdID);
        bool QueryCommandIndeterm(string cmdID);
        string QueryCommandText(string cmdID);
        object QueryCommandValue(string cmdID);
        bool ExecCommand(string cmdID, bool showUI, object value);
        bool ExecCommandShowHelp(string cmdID);
        [return: MarshalAs(UnmanagedType.Interface)]        
        object CreateElement(string eTag);
        void SetOnhelp(object p);
        object GetOnhelp();
        void SetOnclick(object p);
        object GetOnclick();
        void SetOndblclick(object p);
        object GetOndblclick();
        void SetOnkeyup(object p);
        object GetOnkeyup();
        void SetOnkeydown(object p);
        object GetOnkeydown();
        void SetOnkeypress(object p);
        object GetOnkeypress();
        void SetOnmouseup(object p);
        object GetOnmouseup();
        void SetOnmousedown(object p);
        object GetOnmousedown();
        void SetOnmousemove(object p);
        object GetOnmousemove();
        void SetOnmouseout(object p);
        object GetOnmouseout();
        void SetOnmouseover(object p);
        object GetOnmouseover();
        void SetOnreadystatechange(object p);
        object GetOnreadystatechange();
        void SetOnafterupdate(object p);
        object GetOnafterupdate();
        void SetOnrowexit(object p);
        object GetOnrowexit();
        void SetOnrowenter(object p);
        object GetOnrowenter();
        void SetOndragstart(object p);
        object GetOndragstart();
        void SetOnselectstart(object p);
        object GetOnselectstart();
        [return: MarshalAs(UnmanagedType.Interface)]        
        object ElementFromPoint(int x, int y);
        [return: MarshalAs(UnmanagedType.Interface)]        
        object GetParentWindow();
        [return: MarshalAs(UnmanagedType.Interface)]
        object GetStyleSheets();
        void SetOnbeforeupdate(object p);
        object GetOnbeforeupdate();
        void SetOnerrorupdate(object p);
        object GetOnerrorupdate();
        string toString();
        [return: MarshalAs(UnmanagedType.Interface)]
        object CreateStyleSheet(string bstrHref, int lIndex);
    };

    /// <summary>
    /// Provides static methods to assist in programmatic HTML creation
    /// with the System.Windows.Forms.HtmlDocument class
    /// </summary>
    public static class HtmlDocumentFactory
    {
        /// <summary>
        /// Creates a new instance of HtmlDocument class, representing an empty document
        /// </summary>        
        public static HtmlDocument CreateHtmlDocument()
        {            
            return CreateHtmlDocument("");
        }

        /// <summary>
        /// Creates a new instance of HtmlDocument class and initializes its content to
        /// specified string
        /// </summary>        
        public static HtmlDocument CreateHtmlDocument(string content)
        {
            if (content == null) content = "";
            Assembly winforms = typeof(Form).Assembly; //System.Windows.Forms

            //создадим служебный класс HtmlShimManager
            Type t = winforms.GetType("System.Windows.Forms.HtmlShimManager");
            object obj = Activator.CreateInstance(t, true);

            //создадим документ и загрузим в него пустую строку
            var doc = new HTMLDocument();
            IHTMLDocument2 doc2 = (IHTMLDocument2)doc;

            doc2.Write(new object[] { content });
            doc2.Close();

            HtmlDocument htmldoc = null;

            //создаем HtmlDocument с помощью закрытого конструктора
            htmldoc = (HtmlDocument)Activator.CreateInstance(
            typeof(HtmlDocument),
            BindingFlags.Instance | BindingFlags.NonPublic,
            null,
            new object[] { obj, doc },
            System.Globalization.CultureInfo.InvariantCulture);

            return htmldoc;
        }

        /// <summary>
        /// Frees all unmanaged resources associated with HtmlDocument instance
        /// </summary>        
        public static void ReleaseHtmlDocument(HtmlDocument doc)
        {
            if (doc == null) return;
            Type t = typeof(HtmlDocument);
            try
            {
                IDisposable shim = (IDisposable)t.InvokeMember(
                    "shimManager",
                    BindingFlags.GetField | BindingFlags.NonPublic | BindingFlags.Instance,
                    null, doc, new object[0]);
                shim.Dispose();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
            }
            Marshal.ReleaseComObject(doc.DomDocument);
        }

        /// <summary>
        /// Returns the string objet containing all HTML content from HtmlDocument
        /// </summary>        
        public static string HtmlDocumentToString(HtmlDocument doc)
        {
            if (doc == null) throw new ArgumentNullException("Specify HtmlDocument instance");
            var elems = doc.GetElementsByTagName("html");
            if(elems == null)return "";
            if(elems.Count == 0)return "";

            return elems[0].OuterHtml;
        }
    }
}
