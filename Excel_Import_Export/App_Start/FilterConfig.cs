﻿using System.Web;
using System.Web.Mvc;

namespace Excel_Import_Export
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
