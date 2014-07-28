using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Statistic = Microsoft.Office.Interop.Word.WdStatistic;

namespace Word_WritingTracker
{
    public static class Util
    {
        public static bool DEBUG = true;

        /// <summary>
        /// Gets the active document from the application
        /// </summary>
        /// <returns></returns>
        public static Word.Document GetActiveDocumentOrDefault() 
        {
            try
            {
                if (HasAvailableWindows())
                    return Globals.ThisAddIn.Application.ActiveDocument;
                else
                    return default(Word.Document);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                return default(Word.Document);
            }
        }

        /// <summary>
        /// Determines if any word windows are open
        /// </summary>
        /// <returns></returns>
        public static Boolean HasAvailableWindows()
        {
            return Globals.ThisAddIn.Application.Windows.Count > 0;
        }

        /// <summary>
        /// Gets the specified statistic from a given word document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="includeFootnotesAndEndnotes"></param>
        /// <param name="statistic"></param>
        /// <returns></returns>
        public static int GetStatistic(Word.Document document, bool includeFootnotesAndEndnotes, Statistic statistic)
        {
            object include = (object)includeFootnotesAndEndnotes;
            return document.ComputeStatistics(statistic, ref include);
        }

        /// <summary>
        /// Gets the word count from a specified document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="includeFootnotesAndEndnotes"></param>
        /// <returns></returns>
        public static int GetWordCount(Word.Document document, bool includeFootnotesAndEndnotes)
        {
            return GetStatistic(document, includeFootnotesAndEndnotes, Statistic.wdStatisticWords);
        }

        /// <summary>
        /// Gets the page count from a specified document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="includeFootnotesAndEndnotes"></param>
        /// <returns></returns>
        public static int GetPageCount(Word.Document document, bool includeFootnotesAndEndnotes)
        {
            return GetStatistic(document, includeFootnotesAndEndnotes, Statistic.wdStatisticPages);
        }

        /// <summary>
        /// Get the project info (full file name and project name) from a specified document
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static Tuple<String, String> GetProjectInfo(Word.Document document)
        {
            String fullFileName = document.FullName;
            String projectName = System.IO.Path.GetFileNameWithoutExtension(fullFileName);
            return new Tuple<String, String>(fullFileName, projectName);
        }

        public static TrackedFile GetTrackedFile(Tuple<String, String> projectInfo)
        {
            return GetTrackedFile(projectInfo.Item1, projectInfo.Item2);
        }

        public static TrackedFile GetTrackedFile(String fullFilePath, String projectName)
        {
            using (WritingTrackerDataContext db = new WritingTrackerDataContext())
            {
                IEnumerable<TrackedFile> trackedFiles = from tf in db.TrackedFiles
                                                        where tf.ProjectName == projectName
                                                        where tf.FileName == fullFilePath
                                                        select tf;

                return trackedFiles.SingleOrDefault();
            }
        }

        public static TrackedFile GetTrackedFile(String projectName)
        {
            using (WritingTrackerDataContext db = new WritingTrackerDataContext())
            {
                IEnumerable<TrackedFile> trackedFiles = from tf in db.TrackedFiles
                                                        where tf.ProjectName == projectName
                                                        select tf;

                return trackedFiles.SingleOrDefault();
            }
        }

        public static Boolean DocumentIsTracked(Word.Document document)
        {
            Tuple<String, String> projectInfo = GetProjectInfo(document);
            TrackedFile trackedFile = GetTrackedFile(projectInfo);

            if (!trackedFile.IsDefaultForType())
                return trackedFile.Tracked;

            return false;
        }

        public static void InsertTrackedFile(TrackedFile trackedFile)
        {
            using (WritingTrackerDataContext db = new WritingTrackerDataContext())
            {
                db.TrackedFiles.InsertOnSubmit(trackedFile);
                db.SubmitChanges();
            }
        }

        public static void UpdateTrackedFile(TrackedFile trackedFile)
        {
            using (WritingTrackerDataContext db = new WritingTrackerDataContext())
            {
                TrackedFile tracked = (from tf in db.TrackedFiles
                                       where tf.ProjectName == trackedFile.ProjectName
                                       select tf).SingleOrDefault();

                if (!tracked.IsDefaultForType())
                {
                    tracked.FileName = trackedFile.FileName;
                    tracked.Tracked = trackedFile.Tracked;
                    db.SubmitChanges();
                }
            }
        }

        public static void InsertMetric(Word.Document document)
        {
            using (WritingTrackerDataContext db = new WritingTrackerDataContext())
            {
                Tuple<String, String> projectInfo = GetProjectInfo(document);
                TrackedFile tracked = (from tf in db.TrackedFiles
                                       where tf.ProjectName == projectInfo.Item2
                                       select tf).SingleOrDefault();

                if (!tracked.IsDefaultForType())
                {
                    Metric metric = new Metric 
                    {
                        TrackedFile = tracked,
                        WordCount = Util.GetWordCount(document, false),
                        TimeStamp = DateTime.Now
                    };
                    db.Metrics.InsertOnSubmit(metric);
                    db.SubmitChanges();
                }
                
            }
        }

        
    }
}
