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
                // should not reach a state where this occurs (button will be disabled)
                cb.Checked = false;
                MessageBox.Show("Document must be saved before tracking can be applied.","Warning",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // <file path, project name>
            Tuple<String, String> projectInfo = Util.GetProjectInfo(activeDoc);

            TrackedFile dbEntry = Util.GetTrackedFile(projectInfo.Item2);

            if (!dbEntry.IsDefaultForType())
            {
                // check if the file path matches
                if (dbEntry.FileName.Equals(projectInfo.Item1))
                {
                    dbEntry.Tracked = cb.Checked;
                }
                else
                {
                    // prompt user to update the path
                    switch (MessageBox.Show(String.Format("This project name already exists in the database at a different file path.\n\n{0}\n\nDo you want to update the path to match this document?",dbEntry.FileName), "New File Location?", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
                    {
                        case DialogResult.Yes:
                            dbEntry.FileName = projectInfo.Item1;
                            dbEntry.Tracked = cb.Checked;
                            break;
                        default:
                            cb.Checked = false;
                            return;
                    }
                }
                // write changes to database
                Util.UpdateTrackedFile(dbEntry);
            }
            else
            {
                // add a new entry
                TrackedFile entry = new TrackedFile
                {
                    FileName = projectInfo.Item1,
                    Tracked = cb.Checked,
                    ProjectName = projectInfo.Item2
                };

                Util.InsertTrackedFile(entry);
                // insert initial 0 metric
                Util.InsertMetric(activeDoc, 0);
            }
            
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
            ChartForm cf = new ChartForm();
            cf.Show();
        }
        #endregion
    }
}
