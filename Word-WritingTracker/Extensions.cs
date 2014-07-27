using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word_WritingTracker
{
    public static class Extensions
    {
        /// <summary>
        /// Returns the first day of the week of a given datetime
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="startOfWeek"></param>
        /// <returns></returns>
        public static DateTime StartOfWeek(this DateTime dt, DayOfWeek startOfWeek)
        {
            int diff = dt.DayOfWeek - startOfWeek;
            if (diff < 0)
            {
                diff += 7;
            }

            return dt.AddDays(-1 * diff).Date;
        }

        /// <summary>
        /// Returns the first day of the month of a given datetime
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static DateTime StartOfMonth(this DateTime dt)
        {
            return new DateTime(dt.Year, dt.Month, 1);
        }

        /// <summary>
        /// Returns the first day of the quarter of a given datetime
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static DateTime StartOfQuarter(this DateTime dt)
        {
            int quarterNumber = (dt.Month - 1) / 3 + 1;
            return new DateTime(dt.Year, (quarterNumber - 1) * 3 + 1, 1);
        }

        /// <summary>
        /// Returns the first day of the year of a given datetime
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static DateTime StartOfYear(this DateTime dt)
        {
            return new DateTime(dt.Year, 1, 1);
        }
    }
}
