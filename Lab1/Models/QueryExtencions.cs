using System;
using System.Collections.Generic;
using System.EnterpriseServices;
using System.Linq;
using System.Web;

namespace Lab1
{
    public static class QueryExtencions
    {
        public static bool b(this string Val)
        {
            return true;
        }

        public static Dictionary<string, string> Vals(this string val)
        {
            try
            {
                Dictionary<string, string> openWith = new Dictionary<string, string>();
                var items = val.Split(',');
                foreach (var item in items.Where(x=>x!=""))
                {
                    var i = item.Split(':').Where(x => x != "").ToArray();
                    openWith.Add(i[0], i[1]);
                }
                return openWith;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
        }
    }
}