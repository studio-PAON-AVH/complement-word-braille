using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace fr.avh.braille.dictionnaire
{
    public static class Globals
    {
        private static readonly string APPDATA = Path.Combine(
               System.Environment.GetEnvironmentVariable("APPDATA"),
               "AVH",
               "Protection des mots"
        );

        private static readonly string LOGFILE = Path.Combine(AppData.FullName, "protection.log");

        private static StreamWriter logFile = null;

        public static DirectoryInfo AppData { get {
                if(!Directory.Exists(APPDATA)) {
                    return Directory.CreateDirectory(APPDATA);
                } else return new DirectoryInfo(APPDATA);
            }
        }

        #region Gestion des logs
        public delegate void logMessage(string message);
        public delegate void logException(Exception e);

        public static logMessage MessageLoggers;
        public static logException ExceptionLoggers;
        #endregion

        public static void log(string message)
        {
            lock (LOGFILE) {
                MessageLoggers?.Invoke(message);
                if (logFile == null) {
                    logFile = File.AppendText(LOGFILE);
                }
                logFile.WriteLine(message);
                logFile.Flush();
            }
            
        }
        public static void log(Exception e)
        {

            ExceptionLoggers?.Invoke(e);
            lock (LOGFILE) {
                ExceptionLoggers?.Invoke(e);
                if (logFile == null) {
                    logFile = File.AppendText(LOGFILE);
                }
                logFile.WriteLine(e.Message);
                logFile.Write(e.StackTrace);
                logFile.WriteLine();
                logFile.Flush();
            }

        }

        public static async void logAsync(string message)
        {
            await Task.Run(() => log(message));
        }

        public static async void logAsync(Exception e)
        {
            await Task.Run(() => log(e));
        }

    }
}
