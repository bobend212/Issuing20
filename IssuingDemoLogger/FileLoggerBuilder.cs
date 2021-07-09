using System;
using System.Collections.Generic;
using System.Text;

namespace IssuingDemoLogger
{
    public class FileLoggerBuilder : ILoggerBuilder
    {
        private string _path;
        public string LogFormat { get; set; } = "[Message]";
        public ILoggerBuilder WithPath(string path)
        {
            _path = path;
            return this;
        }

        public ILoggerBuilder WithDate()
        {
            LogFormat = "[Date] " + LogFormat;
            return this;
        }

        public ILogger Build()
        {
            return new FileLogger(_path, LogFormat);
        }
    }
}
