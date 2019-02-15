using System;
using System.Collections.Generic;
using System.Configuration;
using System.Threading;


namespace TimeLiveRegression.org.com.utilities
{
    public class ThreadCache1
    {
        private static ThreadCache instance = new ThreadCache();
        ThreadLocal<Dictionary<String, String>> hash = new ThreadLocal<Dictionary<String, String>>();

        public static ThreadCache getInstance()
        {
            return instance;
        }



        public void setProperty(String key, String value)
        {
            this.hash.Value.Add(key, value);
        }

        public String getProperty(String key)
        {
            String returnValue = null;
            if (hash.Value.ContainsKey(key))
            {
                hash.Value.TryGetValue(key, out returnValue);
            }
            return returnValue;
        }

        public void resetProperties()
        {
            hash.Value.Clear();
        }
    }
}
