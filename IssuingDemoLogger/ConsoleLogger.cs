using System;
using System.Collections.Generic;
using System.Text;

namespace IssuingDemoLogger
{
    public class ConsoleLogger : ILogger
    {
        private readonly string _logFormat;

        public ConsoleLogger(string logFormat)
        {
            _logFormat = logFormat;
        }
        public void Log(string message)
        {
            var stringBuilder = new StringBuilder(_logFormat);
            stringBuilder.Replace("[Message]", message);
            stringBuilder.Replace("[Date]", DateTime.Now.ToString());

            Console.WriteLine(stringBuilder.ToString());
        }
    }
}
