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
        public static Word.Document GetActiveDocument() 
        {
            return Globals.ThisAddIn.Application.ActiveDocument;
        }

        public static int GetStatistic(Word.Document document, bool includeFootnotesAndEndnotes, Statistic statistic)
        {
            object include = (object)includeFootnotesAndEndnotes;
            return document.ComputeStatistics(statistic, ref include);
        }

        public static int GetWordCount(Word.Document document, bool includeFootnotesAndEndnotes)
        {
            return GetStatistic(document, includeFootnotesAndEndnotes, Statistic.wdStatisticWords);
        }

        public static int GetPageCount(Word.Document document, bool includeFootnotesAndEndnotes)
        {
            return GetStatistic(document, includeFootnotesAndEndnotes, Statistic.wdStatisticPages);
        }

       
    }
}
