﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TimeLiveRegression.org.com.utilities;

namespace TimeLiveRegression.org.com.screen
{
    public class Note
    {

        public static String sheetname = "Note";
        Common common = CommonManager.getInstance().GetCommon();

        public Boolean SCRNote()
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
