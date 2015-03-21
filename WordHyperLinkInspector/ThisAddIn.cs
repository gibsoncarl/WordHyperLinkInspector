using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace WordHyperLinkInspector
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        public void ValidateHyperlinks()
        {
            Word.Hyperlinks links = this.Application.ActiveDocument.Hyperlinks;
            FailedLinks failedLinks = new FailedLinks();

            foreach (var hyperlink in links)
            {
                var hp = (Word.Hyperlink)hyperlink;

                try
                {
                    WebRequest request = WebRequest.Create(new Uri(hp.Address));
                    request.Method = "HEAD";

                    using (WebResponse response = request.GetResponse())
                    {
                        if (response.ContentLength < 1) failedLinks.Add(hp, "Link content is empty");
                    }
                }
                catch (WebException webEx)
                {
                    string message = webEx.InnerException is FileNotFoundException 
                        ? "File not found" 
                        : "Unable to resolve Hyperlink";
                    
                    failedLinks.Add(hp, message);
                }
            }
        }
    }

    /// <summary>
    /// Dictionary of links and error messages
    /// </summary>
    internal class FailedLinks
    {
        private readonly IDictionary<Word.Hyperlink, string> _links;

        public IDictionary<Word.Hyperlink, string> Links
        {
            get { return new ReadOnlyDictionary<Word.Hyperlink, string>(_links); }
        }

        public FailedLinks()
        {
            _links = new Dictionary<Word.Hyperlink, string>();
        }

        public void Add(Word.Hyperlink link, string message)
        {
            if (_links.ContainsKey(link))
                return;

            _links.Add(link, message);
        }
    }
}
