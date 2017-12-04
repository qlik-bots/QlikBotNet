using System;
using System.IO;

using Telegram.Bot;
using Telegram.Bot.Args;

namespace QlikLog
{
    public class LogFile
    {
        private static string LogFileName;
        private static string LogFilePath;
        public enum LogType { logInfo = 1, logWarning = 2, logError = 3 };

        public LogFile(String path)
        {
            if (path == "" || path == null)
            {
                LogFilePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase.Replace("file:///", "")) + "\\Log";
                System.IO.Directory.CreateDirectory(LogFilePath);
            }
            LogFileName = LogFilePath + "\\QlikSenseBot-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".log";
            AddLine("Starting Log...");
        }

        public string GetLogFileName()
        {
            return LogFileName;
        }

        public void AddLine(string LogLine, bool WriteToConsole = true)
        {
            try
            {
                string line = string.Format("{0}:\t{1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), LogLine);
                using (StreamWriter w = File.AppendText(LogFileName))
                {
                    w.WriteLine(line);
                }
                if (WriteToConsole)
                {
                    Console.WriteLine(line);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("LogFile: {0} Exception caught.", e);
            }
        }

        public void AddBotLine(string LogText, LogType Type = LogType.logInfo)
        {
            string line;
            line = string.Format("{0}\t{1}", Type.ToString("G"), LogText);
            AddLine(line);
        }

        public void AddBotLine(string LogText, string userID, string firstName, string lastName, string username, LogType Type = LogType.logInfo)
        {
            string line;
            line = string.Format("{0}\tID:{1}\tFirst:{2}\tLast:{3}\tUserName:{4}\t<{5}>", Type.ToString("G"), userID, firstName, lastName, username, LogText);
            AddLine(line);
        }

        public void AddBotLine(string LogText, Telegram.Bot.Types.User User, LogType Type = LogType.logInfo)
        {
            string line;
            line = string.Format("{0}\tID:{1}\tFirst:{2}\tLast:{3}\tUserName:{4}\t<{5}>", Type.ToString("G"), User.Id.ToString(), User.FirstName, User.LastName, User.Username, LogText);
            AddLine(line);
        }
    }
}
