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
        /// <summary>
        /// Gets the active document from the application
        /// </summary>
        /// <returns></returns>
        public static Word.Document GetActiveDocument() 
        {
            return Globals.ThisAddIn.Application.ActiveDocument;
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

        
       
    }
}
