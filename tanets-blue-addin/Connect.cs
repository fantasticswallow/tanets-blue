using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Extensibility;
using System.Windows.Forms;
using artfulplace.XmlBlinker;
using System.IO;
using System.Runtime.InteropServices.ComTypes;

namespace tanets_blue_addin
{
    [Guid("850311D5-66DE-478D-857B-D4291C6535E5"), ProgId("TanetsBlue.Connect")]
    public class Connect : Extensibility.IDTExtensibility2, IRibbonExtensibility
    {
        /// <summary>
        /// Fluent UIのカスタムXMLの文字列を返します。
        /// </summary>
        /// <param name="RibbonID"></param>
        /// <returns></returns>
        public string GetCustomUI(string RibbonID)
        {
            return Resource1.String1;
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            
        }

        public void OnBeginShutdown(ref Array custom)
        {
            
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            cli = new tanets_blue.core.OneNoteClient();
            HTMLGenerator.GetBinaryObject = (x1, x2) => cli.GetBinaryObject(x1, x2);
        }

        /// <summary>
        /// アドインがOneNoteから切断された際に呼ばれます
        /// </summary>
        /// <param name="RemoveMode"></param>
        /// <param name="custom"></param>
        /// <remarks>アドインの終了処理を行います</remarks>
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            cli.Dispose();
            this.cli = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void OnStartupComplete(ref Array custom)
        {
            
        }

        private string pathdict = "";

        public void OnMainButtonClick(IRibbonControl control)
        {
            HTMLGenerator.IsResizeFont = IsSizeChange;
            var l = cli.GetCurrentId();
            var px = cli.GetPageContent(l);
            var xe = HTMLGenerator.CreateFromXml(px);
            var id = HTMLGenerator.PageId;
            id = id.Remove(id.LastIndexOf('}')).Remove(0, id.LastIndexOf('{') + 1);
            var xmlPath = System.IO.Path.Combine(pathdict, id + ".html");
            HTMLGenerator.WriteXml(xmlPath, xe);
        }

        private tanets_blue.core.OneNoteClient cli;

        private bool IsSizeChange = false;

        public void OnSizeCheckBoxClick(IRibbonControl control, bool pressed)
        {
            IsSizeChange = pressed;
        }

        public IStream OnGetImage(string imageName)
        {
            MemoryStream stream = new MemoryStream();
            
            // TODO : Write image data to stream.

            return new ReadOnlyIStreamWrapper(stream);
        }
    }
}
