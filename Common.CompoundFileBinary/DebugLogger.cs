using System;
using System.Diagnostics;
using Microsoft.Extensions.Logging;

namespace b2xtranslator.CompoundFileBinary
{
    public class DebugLogger : ILogger
    {
        private readonly string _categoryName;

        public static DebugLogger Default { get; } = new DebugLogger("Default");


        public DebugLogger(string categoryName)
        {
            _categoryName = categoryName;
        }

        public IDisposable BeginScope<TState>(TState state) => NullScope.Instance;

        public bool IsEnabled(LogLevel logLevel) => true;

        public void Log<TState>(
            LogLevel logLevel,
            EventId eventId,
            TState state,
            Exception exception,
            Func<TState, Exception?, string> formatter)
        {
            if (formatter == null) return;

            string message = formatter(state, exception);
            Debug.WriteLine($"[{logLevel}] {_categoryName}: {message}");

            if (exception != null)
            {
                Debug.WriteLine(exception);
            }
        }

        private class NullScope : IDisposable
        {
            public static NullScope Instance { get; } = new NullScope();
            public void Dispose() { }
        }
    }
}