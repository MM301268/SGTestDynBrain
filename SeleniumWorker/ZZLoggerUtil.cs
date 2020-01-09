using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumWorker
{
    public static class ZZLoggerUtil
    {
        #region CreatUnderlines
        public static string CreateUnderlines(string argString)
        {
            var underlines = string.Empty;
            var loglinelength = argString.Length;
            for (var i = 0; i < loglinelength; i++)
            {
                underlines += '=';
            }
            return underlines;
        }
        #endregion
    }
}
