using System;
using System.Security.AccessControl;
using System.Security.Principal;
using System.IO;
using System.Text;

namespace bns.spo.diagnostic
{
    public class LogManager
    {
                           
        #region helpers
        public static string GetFileOwner(string file)
        {
            string owner = "not found";
            try
            {

                FileSecurity fileSecurity = new FileSecurity(file, AccessControlSections.Owner);                
                IdentityReference identityReference = fileSecurity.GetOwner(typeof(SecurityIdentifier));
                owner = identityReference.Value;

                NTAccount nt = identityReference.Translate(typeof(NTAccount)) as NTAccount;
                owner = nt.Value;
            }
            catch 
            {
               
            }
            return owner;
        }              
        #endregion

    }

    public class LogCsvFile
    {
        public string Header;
        public string Name;
        private string _reportpath;
        private string _reportid;

        public LogCsvFile(string header, string name)
        {
            Header = header;
            Name = name;                      
            _reportid = Guid.NewGuid().ToString();
            _reportpath = string.Format("{0}_{1}.csv", Name, _reportid); 
        }



        public string ReportPath
        {
            get
            {

                return _reportpath;
            }
            set
            {
                _reportpath = value;
            }
        }
        private string StatusReportPath
        {
            get
            {
                string p = "{0}_{1}.txt";
                return string.Format(p, Name, _reportid);
            }
        }

        public void WriteToCVS(params string[] content)
        {
            StringBuilder s = new StringBuilder();
            s.AppendLine(ParseContent(content));
            if (!File.Exists(ReportPath))
                using (StreamWriter sw = new StreamWriter(ReportPath))
                    sw.WriteLine(Header);
            File.AppendAllText(ReportPath, s.ToString());
        }
        public void WriteInfo(string text)
        {
            string s = "[info]:" + text;
            WriteStatusToTxt(s);
        }
        public void WriteError(string text)
        {
            string s = "[error]:" + text;
            WriteStatusToTxt(s);
        }
        public void WriteWarning(string text)
        {
            string s = "[warning]: " + text;
            WriteStatusToTxt(s);
        }

        private void WriteStatusToTxt(string text)
        {
            try
            {
                StringBuilder s = new StringBuilder();
                s.AppendLine(text);
                if (!File.Exists(StatusReportPath))
                    using (FileStream f = File.Create(StatusReportPath))
                        f.Close();
                File.AppendAllText(StatusReportPath, s.ToString());
                Console.WriteLine(s.ToString());
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        private string ParseContent(params string[] contentfields)
        {
            string r = string.Empty;
            foreach (string c in contentfields)
                r += "," + c.Replace(",", "");
            return r.Substring(1);
        }
    }
}
