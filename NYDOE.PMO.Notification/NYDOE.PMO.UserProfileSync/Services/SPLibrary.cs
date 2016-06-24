using System;
using System.Configuration;
using System.IO;

namespace NYDOE.PMO.UserProfileSync.Services
{
    public static class SPLibrary
    {
        public static void WriteErrorLog(Exception ex)
        {
            StreamWriter sw = null;
            try
            {
                string logfile = String.Format("{0}\\{1}",AppDomain.CurrentDomain.BaseDirectory,ConfigurationManager.AppSettings["logfile"].ToString());
                sw = new StreamWriter(logfile, true);
                sw.WriteLine(String.Format("{0} : {1} - {2}", DateTime.Now.ToString(), ex.Source.ToString().Trim(), ex.Message.ToString()));
                sw.Flush();
                sw.Close();
            }
            catch
            {

            }

        }
        public static void WriteToFile(string text)
        {            
            using (StreamWriter writer = new StreamWriter(ConfigurationManager.AppSettings["logfile"].ToString(), true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));                
                writer.Close();
            }
        }

    }
}
