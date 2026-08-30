using System;
using System.IO;

public static class Logger
{
    private static readonly string LogDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        @"WorkOnOff\logs"
    );

    static Logger()
    {
        try
        {
            if (!Directory.Exists(LogDirectory))
            {
                Directory.CreateDirectory(LogDirectory);
            }
            CleanupOldLogs();
        }
        catch {}
    }

    public static void Write(string message, string level = "INFO")
    {
        try
        {
            string dateStr = DateTime.Now.ToString("yyyyMMdd");
            string logPath = Path.Combine(LogDirectory, string.Format("app_{0}.log", dateStr));
            string logTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string logLine = string.Format("[{0}] [{1}] {2}{3}", logTime, level, message, Environment.NewLine);

            File.AppendAllText(logPath, logLine);
        }
        catch {}
    }

    private static void CleanupOldLogs()
    {
        try
        {
            if (!Directory.Exists(LogDirectory)) return;

            string[] files = Directory.GetFiles(LogDirectory, "app_*.log");
            DateTime cutoff = DateTime.Now.AddDays(-7);

            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.LastWriteTime < cutoff)
                {
                    fi.Delete();
                }
            }
        }
        catch {}
    }
}
