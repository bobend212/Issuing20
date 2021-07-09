using System;

namespace IssuingDemoLogger
{
    class Program
    {
        static void Main(string[] args)
        {
            var logger = new FileLoggerBuilder().WithPath("logs.txt").WithDate().Build();

            logger.Log("Hello World");
        }
    }
}
