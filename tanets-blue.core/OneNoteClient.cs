using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.OneNote;

namespace tanets_blue.core
{
    public class OneNoteClient
    {
        public OneNoteClient()
        {
            app = new ApplicationClass();
        }

        internal ApplicationClass app { get; set; }


        public string GetCurrentId()
        {
            return app.Windows.CurrentWindow.CurrentPageId;
        }

        public string GetPageContent(string id)
        {
            var res = "";
            app.GetPageContent(id, out res);
            return res;

        }

        public void UpdatePageContent(string str)
        {
            app.UpdatePageContent(str);
        }

        public string GetBinaryObject(string pid, string clid)
        {
            var res = "";
            app.GetBinaryPageContent(pid, clid, out res);
            return res;
        }

        public void Dispose()
        {
            app = null;
        }

    }
}
