using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ChatExcel.Addin.Utilties
{
    public static class Log
    {
        private static string _logFileDir;
        private const string DateTimeStampFormat = "yyyy-MM-dd HH:mm:ss.fff";

        #region 私有属性
        // 如果日志文件名为空，则每日自动生成log
        private static StreamWriter LogFile
        {
            get
            {
                if (_currentFileStream != null && DateTime.Now.Day == _currentDay) return _currentFileStream;

                // 创建当日log文件
                _currentDay = DateTime.Now.Day;

                //如果不显性指定，默认在TEMP目录中创建Logs文件目录
                if (string.IsNullOrEmpty(LogFileDir))
                {
                    //GetTempPath获取临时文件夹的路径，以反斜杠结尾。类似 C:\Users\UserName\AppData\Local\Temp\  
                    LogFileDir = Path.Combine(Path.GetTempPath(), "logs");
                }
                try
                {
                    if (!Directory.Exists(LogFileDir)) Directory.CreateDirectory(LogFileDir);
                }
                catch
                {
                    LogFileDir = null;
                }
                if (!string.IsNullOrEmpty(LogFileDir))
                {
                    // log全路径及文件名
                    string logFile = Path.Combine(LogFileDir, DateTime.Now.ToString(@"yyMMdd.lo\g"));

                    if (_currentFileStream != null) _currentFileStream.Dispose();

                    try
                    {
                        _currentFileStream = new StreamWriter(File.Open(logFile, FileMode.Append, FileAccess.Write, FileShare.ReadWrite));
                        _currentFileStream.AutoFlush = true;

                        return _currentFileStream;
                    }
                    catch
                    {
                        return null;
                    }
                }
                else
                    return null;
            }
        }
        #endregion 私有属性

        #region 公共属性
        public static string TimeStampFormat { get; set; } = "HH:mm:ss.fff";
        public static bool FileOutputEnabled { get; set; } = true;  // 是否允许文件输出
        public static bool ConsoleOutputEnabled { get; set; } = false;
        public static bool InfoLogEnabled { get; set; } = true;
        public static bool ErrorLogEnabled { get; set; } = true;
        public static bool TraceLogEnabled { get; set; } = true;

        public static string LogFileDir
        {
            get
            {
                return _logFileDir;
            }
            set
            {
                if (string.IsNullOrEmpty(value)) return;
                _logFileDir = value;
                if (!_logFileDir.EndsWith(Path.DirectorySeparatorChar.ToString()))
                {
                    _logFileDir += Path.DirectorySeparatorChar;
                }
            }
        }
        #endregion 公共属性

        /// <summary>
        /// 最好订阅此方法以“退出”应用程序的事件。但并非所有情况都可以覆盖。
        /// 没有特定于应用程序的退出事件。
        /// </summary>
        public static void DoGracefullExit()
        {
            string msg = string.Format("-------- {0} Logging system is shutdown --------\r\n", DateTime.Now.ToString(DateTimeStampFormat));
            LogBlockingQueue.Add(msg);
            try
            {
                LogBlockingQueue.CompleteAdding();
                _writingThread.Join();
            }
            catch
            { }
        }

        #region 条件编译输出控制
        [Conditional("FAST_LOG")]
        public static void Info(string msg, params object[] par)
        {
            if (!InfoLogEnabled || LogBlockingQueue.IsAddingCompleted) return;
            LogBlockingQueue.Add(string.Format("{0} INFO {1}\r\n",
                                   DateTime.Now.ToString(TimeStampFormat),
                                   (par == null || par.Length == 0) ? msg : string.Format(msg, par)));  // 注意：{}符号可能会破坏格式化调用
        }
        [Conditional("FAST_LOG")]
        public static void Warn(string msg, params object[] par)
        {
            if (!InfoLogEnabled || LogBlockingQueue.IsAddingCompleted) return;
            LogBlockingQueue.Add(string.Format("{0} !!!WARNING!!! {1}\r\n",
                                   DateTime.Now.ToString(TimeStampFormat),
                                   (par == null || par.Length == 0) ? msg : string.Format(msg, par)));  // 注意：{}符号可能会破坏格式化调用
        }

        [Conditional("FAST_LOG")]
        public static void Error(string msg, params object[] par)
        {
            if (!ErrorLogEnabled || LogBlockingQueue.IsAddingCompleted) return;
            LogBlockingQueue.Add(string.Format("{0} ***ERROR*** {1}\r\n",
                                   DateTime.Now.ToString(TimeStampFormat),
                                   (par == null || par.Length == 0) ? msg : string.Format(msg, par)));
        }

        [Conditional("FAST_LOG")]
        public static void Error(Exception ex)
        {
            Error(ex.ToString());
        }

        //https://www.jianshu.com/p/76e9005763d0 在 .NET 4.0 中使用 .NET 4.5 中新增的特性
        [AttributeUsage(AttributeTargets.Parameter, Inherited = false)]
        public class CallerMemberNameAttribute : Attribute { }

        [AttributeUsage(AttributeTargets.Parameter, Inherited = false)]
        public class CallerFilePathAttribute : Attribute { }

        [AttributeUsage(AttributeTargets.Parameter, Inherited = false)]
        public class CallerLineNumberAttribute : Attribute { }

        [Conditional("FAST_TRACE")]
        public static void Trace(string msg, [CallerMemberName] string memberName = "",
                                                [CallerFilePath] string sourceFilePath = "",
                                                [CallerLineNumber] int sourceLineNumber = 0)
        {
            if (!TraceLogEnabled || LogBlockingQueue.IsAddingCompleted) return;

            LogBlockingQueue.Add(string.Format("{0} TRC {1}:{2}[{3}] {4}\r\n", DateTime.Now.ToString(TimeStampFormat),
                                                                                      Path.GetFileName(sourceFilePath), memberName,
                                                                                      sourceLineNumber, msg));
        }
        #endregion 条件编译输出控制

        // 私有变量
        private static readonly Thread _writingThread;
        private static readonly BlockingCollection<string> LogBlockingQueue;
        private static int _currentDay = -1;
        private static StreamWriter _currentFileStream;

        // 静态构造函数
        static Log()
        {
            LogBlockingQueue = new BlockingCollection<string>();
            //AppDomain.CurrentDomain.ProcessExit += (s, a) => DoGracefullExit();     //不一定能执行

            //开启后台线程写入日志
            _writingThread = new Thread(Write);
            _writingThread.Name = "Log writer thread";
            _writingThread.IsBackground = true;                              //后台线程，跟随主线程退出。
            _writingThread.Priority = ThreadPriority.BelowNormal;
            _writingThread.Start();

            string msg = string.Format("-------- {0} Logging system is initialized --------\r\n", DateTime.Now.ToString(DateTimeStampFormat));
            LogBlockingQueue.Add(msg);
        }

        // 日志记录，无限循环运行
        private static void Write()
        {
            string msg;
            try
            {
                while (true)
                {
                    if (LogBlockingQueue.IsCompleted) break;
                    msg = LogBlockingQueue.Take();

                    if (LogFile != null)
                    {
                        if (FileOutputEnabled) LogFile.Write(msg);

                        if (ConsoleOutputEnabled)
                        {
                            ConsoleColor oldColor = Console.ForegroundColor;
                            string type = msg.Substring(9, 3);
                            switch (type)
                            {
                                case "ERR":
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    break;
                                case "INF":
                                    Console.ForegroundColor = ConsoleColor.Cyan;
                                    break;
                            }
                            Console.Write(msg);
                            Console.ForegroundColor = oldColor;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                msg = string.Format("-------- {0} Log system have an ERROR --------\r\n", DateTime.Now.ToString(DateTimeStampFormat));
                msg += ex.Message + "\r\n";
                msg += "---------------------------------------------------------------------------\r\n";
                if (FileOutputEnabled && LogFile != null) LogFile.Write(msg);
                if (ConsoleOutputEnabled) Console.Write(msg);
            }
        }

        public static void flush()
        {
            if (FileOutputEnabled && LogFile != null) LogFile.Flush();
        }
    }
}
