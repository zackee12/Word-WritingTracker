using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace Word_WritingTracker
{
    public partial class HomeRibbon
    {
        private void HomeRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            System.Diagnostics.Debug.WriteLineIf(Util.DEBUG, "HomeRibbon_Load");
        }

        #region Control_Events
        private void checkBoxTrackMetrics_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox cb = sender as RibbonCheckBox;
           
            Microsoft.Office.Interop.Word.Document activeDoc = Util.GetActiveDocumentOrDefault();
            
            // check if the active doc was found and that it has a path
            if (activeDoc.IsDefaultForType())
            {
                cb.Checked = false;
                return;
            }
            else if (String.IsNullOrEmpty(activeDoc.Path)){
                cb.Checked = false;
                MessageBox.Show("Document must be saved before tracking can be applied.","Warning",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // <file path, project name>
            Tuple<String, String> projectInfo = Util.GetProjectInfo(activeDoc);

            
        }

        private void buttonExport_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Sorry! :(  This feature hasn't been implemented yet.", "Not Implemented", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void buttonSettings_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Sorry! :(  This feature hasn't been implemented yet.", "Not Implemented", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void buttonChart_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Sorry! :(  This feature hasn't been implemented yet.", "Not Implemented", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        #endregion
    }
}
