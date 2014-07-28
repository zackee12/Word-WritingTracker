using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace Word_WritingTracker
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Diagnostics.Debug.WriteLineIf(Util.DEBUG, "ThisAddIn_Startup");
            Word.Application app = Globals.ThisAddIn.Application;
            app.DocumentChange += app_DocumentChange;
            app.DocumentBeforeSave += app_DocumentBeforeSave;
            app.DocumentBeforeClose += app_DocumentBeforeClose;
            app.DocumentOpen += app_DocumentOpen;
            
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            System.Diagnostics.Debug.WriteLineIf(Util.DEBUG, "ThisAddIn_Shutdown");
        }

        void app_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            System.Diagnostics.Debug.WriteLineIf(Util.DEBUG, "DocumentBeforeClose");
        }

        void app_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            System.Diagnostics.Debug.WriteLineIf(Util.DEBUG, "DocumentBeforeSave");

            app_DocumentChange();

            Microsoft.Office.Tools.Ribbon.RibbonCheckBox cb = Globals.Ribbons.HomeRibbon.checkBoxTrackMetrics;
            Word.Document activeDoc = Util.GetActiveDocumentOrDefault();
            
            if (cb.Checked && !activeDoc.IsDefaultForType() && Util.DocumentIsTracked(activeDoc))
            {
                Util.InsertMetric(activeDoc);
            }
        }

        void app_DocumentChange()
        {
            System.Diagnostics.Debug.WriteLineIf(Util.DEBUG, "DocumentChange");
            
            Microsoft.Office.Tools.Ribbon.RibbonCheckBox cb = Globals.Ribbons.HomeRibbon.checkBoxTrackMetrics;
            Word.Document activeDoc = Util.GetActiveDocumentOrDefault();
            
            if (!activeDoc.IsDefaultForType() && !String.IsNullOrEmpty(activeDoc.Path))
            {
                cb.Enabled = true;
                cb.Checked = Util.DocumentIsTracked(activeDoc);
            }
            else
            {
                // disable the button if the active doc or path doesn't exist
                cb.Enabled = false;
                cb.Checked = false;
            }
                
        }

        void app_DocumentOpen(Word.Document Doc)
        {
            System.Diagnostics.Debug.WriteLineIf(Util.DEBUG, "DocumentOpen");
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
    }
}
