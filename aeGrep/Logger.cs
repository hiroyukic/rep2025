using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;

//;;;;;;;;;;;;;;;;;;;;;;;;

namespace aeGrep
{
    class Logger
    {
        static string LOGDIR_PATH = "KK";

        //<summary> ログレベル<summary>
        private enum LogLevel1
        {
            ERROR,
            WARN,
            INFO,
            DEBUG
        }

        private static Logger singleton = null;
        //ログファイルのフルパス
        public readonly string logFilePath = null;
        private readonly object lockObj = new object();
        private StreamWriter Stream = null;


        //<summary>インスタンス生成<summary>
        public static Logger GetInstance(string logFileName)
        {
            return new Logger(logFileName);
        }

        //<summary>コンストラクタ<summary>
        private Logger(string logFileName)
        {
            string LOGFILE_NAME = "LOG_" + logFileName + ".csv";
            try
            {
                //iniファイルのパスを作成
                this.logFilePath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\\" + LOGFILE_NAME;

                //古いログファイルを削除
                 File.Delete(logFilePath);

                 //ログファイルを生成
                 CreateLogfile(new FileInfo(logFilePath));
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR " + ex.Message);
                MessageBox.Show(ex.Message + Environment.NewLine + "アプリケーションを終了します。" + "Error");
            }
        }




        //<summary>ログを出力する<summary>
        private void Out(LogLevel1 level, string msg)
        {
            int tid = System.Threading.Thread.CurrentThread.ManagedThreadId;
            string fullMsg = string.Format("[{0}][{1}][{2}][{3}]", DateTime.Now.ToString("yyyy-MM-dd HH:mm"), tid, level.ToString(), msg);

            //lock (this.lockObj)
            {
                this.Stream.WriteLine(fullMsg);
            }

        }


        //<summary>ログを出力する<summary>
        public void OutLine(int level, string msg)
        {
            int tid = System.Threading.Thread.CurrentThread.ManagedThreadId;
            string fullMsg = string.Format("[{0}][{1}][{2}][{3}]", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), tid, level.ToString(), msg);

            //lock (this.lockObj)
            {
                this.Stream.WriteLine(fullMsg);
            }

        }













        //<summary>ログファイルを生成<summary>
        private void  CreateLogfile(FileInfo  logFile)
        {
            if (Directory.Exists(logFile.DirectoryName))
            {
                Directory.CreateDirectory(logFile.DirectoryName);
            }
            this.Stream = new StreamWriter(logFile.FullName, true, Encoding.UTF8)
            { AutoFlush = true };
        }


        //<summary>close<summary>
        public void Close()
        {
            this.Stream.Close();
        }




    }
}
