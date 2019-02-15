﻿using System;
using TimeLiveRegression.org.com.utilities;

namespace TimeLiveRegression.org.com.screen
{
    public class SearchHotel
    {

        public static String sheetname = "SearchHotel";
        Common common = CommonManager.getInstance().GetCommon();

        public Boolean SCRSearchHotel()
        {
            Boolean status = true;
            status = common.RunComponent(sheetname, Common.o);
            if (!status)
            {
                status = false;
            }
            return status;
        }

    }
}
