using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TimeLiveRegression.org.com.utilities;
namespace TimeLiveRegression.org.com.utilities
{
    public class CommonManager
    {

        private static CommonManager instance = new CommonManager();
        ThreadLocal<Common> common = new ThreadLocal<Common>();

        public static CommonManager getInstance()
        {
            return instance;
        }

        public void SetCommon(Common com)
       {
            this.common.Value = com;
        }

        public Common GetCommon()
        {
            return common.Value;

        }

    }
}
