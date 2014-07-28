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
            InsertMetric(document, Util.GetWordCount(document, false));
        }

        public static void InsertMetric(Word.Document document, int wordCount)
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
                        WordCount = wordCount,
                        TimeStamp = DateTime.Now
                    };
                    db.Metrics.InsertOnSubmit(metric);
                    db.SubmitChanges();
                }

            }
        }

        public static Dictionary<TrackedFile, List<Metric>> GetLastMetricOfDayForTrackedProjects()
        {
            var dict = new Dictionary<TrackedFile, List<Metric>>();
            var db = new WritingTrackerDataContext();

            IEnumerable<TrackedFile> currentTrackedQuery = from t in db.TrackedFiles
                                                           where t.Tracked == true
                                                           select t;

            foreach (TrackedFile tracked in currentTrackedQuery)
            {
                List<Metric> lastMetricEachDay = (from metric in tracked.Metrics
                                                  let time = metric.TimeStamp
                                                  group metric by new { timestamp = time.Date } into g
                                                  select g.OrderByDescending(t => t.TimeStamp).FirstOrDefault()).ToList();

                dict.Add(tracked, lastMetricEachDay);
            }
            return dict;
        }

        public static Dictionary<TrackedFile, List<Metric>> GetLastMetricOfDayForTrackedProjects(DateTime startDate, DateTime endDate)
        {
            var dict = new Dictionary<TrackedFile, List<Metric>>();
            var db = new WritingTrackerDataContext();

            IEnumerable<TrackedFile> currentTrackedQuery = from t in db.TrackedFiles
                                                           where t.Tracked == true
                                                           select t;

            foreach (TrackedFile tracked in currentTrackedQuery)
            {
                
                List<Metric> lastMetricEachDay = (from metric in tracked.Metrics
                                                        let time = metric.TimeStamp
                                                        where time.Date >= startDate.Date
                                                        where time.Date <= endDate.Date
                                                        group metric by new { timestamp = time.Date } into g
                                                        select g.OrderByDescending(t => t.TimeStamp).FirstOrDefault()).ToList();
                dict.Add(tracked, lastMetricEachDay);
            }
            return dict;
        }

        public static Dictionary<String, List<Tuple<DateTime, int>>> GetDailyWordCount(DateTime startDate, DateTime endDate)
        {
            var dict = GetLastMetricOfDayForTrackedProjects();
            var wordDict = new Dictionary<String, List<Tuple<DateTime, int>>>();
            foreach (TrackedFile tf in dict.Keys)
            {
                var metricList = new List<Metric>();
                if (!dict.TryGetValue(tf, out metricList))
                    System.Diagnostics.Debug.WriteLine("Failed to get metricList from dictionary");

                metricList = metricList.OrderByDescending(m => m.TimeStamp).ToList();
                //DateTime start = metricList.Last().Timestamp;
                //DateTime end = metricList.First().Timestamp;
                DateTime start = startDate;
                DateTime end = endDate;

                var list = new List<Tuple<DateTime, int>>();
                // fill in each date delta word count
                for (DateTime date = end.Date; date >= start.Date; date = date.AddDays(-1))
                {
                    // get the current date
                    Metric current = metricList.SingleOrDefault(m => m.TimeStamp.Date == date);
                    // get the date before
                    Metric next = metricList.FirstOrDefault(m => m.TimeStamp.Date < date);

                    int wordDelta;
                    if (current != null && next != null)
                        wordDelta = current.WordCount - next.WordCount;
                    else
                        wordDelta = 0;

                    list.Add(new Tuple<DateTime, int>(date, wordDelta));
                }
                wordDict.Add(tf.ProjectName, list);
            }
            return wordDict;
        }
        
    }
}
