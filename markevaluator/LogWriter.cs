using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;

namespace markevaluator
{
    class LogWriter
    {
        static String directory_path;
        static String file_name;
        static FileStream fStream;
        static String temp;
        static byte[] content;

        /// <summary>
        /// Writes errors in log file
        /// </summary>
        /// <param name="task">on which operation</param>
        /// <param name="message">exception message</param>
        public static void WriteError(String task,String message)
        {
            try
            {
                if (!Directory.Exists(Environment.CurrentDirectory + "\\Logs"))
                    Directory.CreateDirectory("Logs");

                directory_path = Environment.CurrentDirectory + "\\Logs\\";
                file_name = "Error_log_" + DateTime.Today.ToShortDateString().Replace('/', '-') + ".txt";

                fStream = File.Open(directory_path + file_name, FileMode.Append, FileAccess.Write, FileShare.Read);

                temp = "[" + DateTime.Now.ToString() + "]" + Environment.NewLine;
                temp += "Task: " + task + Environment.NewLine + "Error Info: " + message + Environment.NewLine;
                temp += "--------------------------------------------------------------------------------------" + Environment.NewLine;
                content = Encoding.ASCII.GetBytes(temp);

                fStream.Write(content, 0, content.Length);
                fStream.Close();
            }
            catch(Exception)
            {
                //LOL
            }
        }
    }
}
