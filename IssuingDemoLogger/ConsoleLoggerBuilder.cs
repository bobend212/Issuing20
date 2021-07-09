using System;
using System.Collections.Generic;
using System.Text;

namespace IssuingDemoLogger
{
    public class ConsoleLoggerBuilder : ILoggerBuilder
    {
        public string LogFormat { get; set; } = "[Message]";
        public ILoggerBuilder WithPath(string path)
        {
            return this;
        }

        public ILoggerBuilder WithDate()
        {
            LogFormat = "[Date] " + LogFormat;
            return this;
        }

        public ILogger Build()
        {
            return new ConsoleLogger(LogFormat);
        }
    }
}
