using System;
using System.Collections.Generic;
using System.Text;

namespace IssuingDemoLogger
{
    public interface ILogger
    {
        void Log(string message);
    }
}
