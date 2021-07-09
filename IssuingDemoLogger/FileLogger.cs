using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace IssuingDemoLogger
{
    public class FileLogger : ILogger
    {
        private readonly string _path;
        private readonly string _logFormat;

        public FileLogger(string path, string logFormat)
        {
            _path = path;
            _logFormat = logFormat;
        }

        public void Log(string message)
        {
            var stringBuilder = new StringBuilder(_logFormat);
            stringBuilder.Replace("[Message]", message);
            stringBuilder.Replace("[Date]", DateTime.Now.ToString());

            File.AppendAllText(_path, Environment.NewLine);
            File.AppendAllText(_path, stringBuilder.ToString());
        }
    }
}
