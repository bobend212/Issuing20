using System;
using System.Collections.Generic;
using System.Text;

namespace IssuingDemoLogger
{
    public interface ILoggerBuilder
    {
        string LogFormat { get; set; }
        ILoggerBuilder WithPath(string path);
        ILoggerBuilder WithDate();
        ILogger Build();

    }
}
