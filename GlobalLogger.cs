using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
using log4net;
using log4net.Repository.Hierarchy;
using log4net.Core;
using log4net.Appender;
using log4net.Layout;

namespace GlobalLibrary
{
    public class GlobalLogger
    {
        private PatternLayout _layout = new PatternLayout();
        private const string LOG_PATTERN = "%date{yyyy-MM-dd HH:mm:ss.fff} | %-5level | %message%newline";

        public string DefaultPattern
        {
            get { return LOG_PATTERN; }
        }

        public GlobalLogger()
        {
            _layout.ConversionPattern = DefaultPattern;
            _layout.ActivateOptions();
        }

        public void AddAppender(IAppender appender)
        {
            Hierarchy hierarchy = (Hierarchy)LogManager.GetRepository();
            hierarchy.Root.AddAppender(appender);
        }

        static GlobalLogger()
        {
            Hierarchy hierarchy = (Hierarchy)LogManager.GetRepository();
            TraceAppender tracer = new TraceAppender();
            PatternLayout patternLayout = new PatternLayout();

            patternLayout.ConversionPattern = LOG_PATTERN;
            patternLayout.ActivateOptions();

            tracer.Layout = patternLayout;
            tracer.ActivateOptions();
            hierarchy.Root.AddAppender(tracer);

            RollingFileAppender roller = new RollingFileAppender();
            roller.Layout = patternLayout;
            roller.AppendToFile = true;
            roller.RollingStyle = RollingFileAppender.RollingMode.Size;
            roller.MaxSizeRollBackups = 4;
            roller.MaximumFileSize = "10480KB";
            roller.File = @"c:\Logs\Automation.txt";
            roller.ActivateOptions();
            hierarchy.Root.AddAppender(roller);

            hierarchy.Root.Level = Level.All;
            hierarchy.Configured = true;
        }

        public static ILog create()
        {
            // For writing on Console - Comment this if writing to conslole is not needed
            log4net.Config.BasicConfigurator.Configure();

            return LogManager.GetLogger("Automation");
        }
    }

    public class Log
    {
        public static readonly ILog log = GlobalLibrary.GlobalLogger.create();

        public static void Error(string msg)
        {
            log.Error(msg);
        }

        public static void Error(Exception ex)
        {
            log.Error(ex.Message + "\n\n" + ex.StackTrace);
            Debug();
        }

        public static void Info(string msg)
        {
            log.Info(msg);
        }

        public static void Fatal(string msg)
        {
            log.Fatal(msg);
        }

        public static void Debug()
        {
            //Print the CALL STACK
            string debugTrace = new System.Diagnostics.StackTrace().ToString();
            int index = debugTrace.IndexOf("at System.RuntimeMethodHandle._InvokeMethodFast");

            if (index > 0)
            {
                log.Debug(new System.Diagnostics.StackTrace().ToString().Substring(0, index));
            }
            else
            {
                log.Debug(debugTrace);
            }
        }

        public static void Debug(string msg)
        {
            log.Debug(msg);
        }

        public static void DebugFalseReturn(string msg)
        {
            log.Debug(msg + " - Method returned false");
        }

        public static void Warn(string msg)
        {
            log.Warn(msg);
        }

    }
}
