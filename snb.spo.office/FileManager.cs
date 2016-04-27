using System;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using bns.spo.diagnostic;
using System.Configuration;


namespace bns.spo.files
{
    class FileManager
    {

        #region helpers

        public static MigrationSettigs Init(string doTransit, string from_lastdate)
        {
            MigrationSettigs d = new MigrationSettigs();
        
            d.LastDate = from_lastdate;           
            d.ExcludeFolders = ConfigurationManager.AppSettings["migration.folders.exclude"];

            d.SearchPattern = ConfigurationManager.AppSettings["user.folder.pattern"];
            if (doTransit == "1")
            {
                d.SearchPattern = ConfigurationManager.AppSettings["transit.folder.pattern"];                           
            }

            d.JobsSourcePsPathKeyValue = ConfigurationManager.AppSettings["migration.source.path"]; 
            d.JobsSourcePsDisplayUrlKeyValue = ConfigurationManager.AppSettings["migration.source.displayurl"];
            d.JobsSourcePsUrlKeyValue = ConfigurationManager.AppSettings["migration.source.url"];

            d.JobsTargetPsPathKeyValue = ConfigurationManager.AppSettings["migration.target.path"];
            d.JobsTargetPsDisplayUrlKeyValue = ConfigurationManager.AppSettings["migration.target.displayurl"];
            d.JobsTargetPsUrlKeyValue = ConfigurationManager.AppSettings["migration.target.url"];
            
            return d;
        }

        #endregion


        #region Delegates

        public delegate void ProcessFileDelegate(string parent, string x, MigrationSettigs d, LogCsvFile log);
        public static void ProcessFileValidations(string parent, string x, MigrationSettigs d, LogCsvFile log)
        {
            log.Header = "Folder,File,Length,ErrorType,Owner";            

            try
            {  //# % * : < > ? / \ |    
                if (x.IndexOf("~")!=-1 
                    || x.IndexOf("#") != -1 
                    || x.IndexOf("%") != -1 
                    || x.IndexOf("*") != -1                   
                    || x.IndexOf("<") != -1
                    || x.IndexOf(">") != -1
                    || x.IndexOf("?") != -1
                    || x.IndexOf("|") != -1)
                {
                    log.WriteToCVS(parent, x, x.Length.ToString(), "Invalid Char", LogManager.GetFileOwner(x));
                }
                              
                if (x.Length > 260)
                {                   
                    log.WriteToCVS(parent, x, x.Length.ToString(), "URL Length", LogManager.GetFileOwner(x));
                }
                else
                {
                    string[] parts = x.Split(new char[] { '\\' });
                    foreach (string part in parts)
                    {
                        if (part.Length > 128)
                        {
                            log.WriteToCVS(parent, x, x.Length.ToString(), "Name Length", LogManager.GetFileOwner(x));
                            break;
                        }

                        if ((part.IndexOf(".") == -1)  && 
                            (part.EndsWith("_file", StringComparison.CurrentCultureIgnoreCase) ||
                             part.EndsWith("_files", StringComparison.CurrentCultureIgnoreCase) ||
                             part.EndsWith("-filer", StringComparison.CurrentCultureIgnoreCase) ||
                             part.EndsWith("_fails", StringComparison.CurrentCultureIgnoreCase)))
                        {
                            log.WriteToCVS(parent, x, x.Length.ToString(), "Invalid Name", LogManager.GetFileOwner(x));
                            break;
                        }                    
                    }                   
                }

            }
            catch (Exception ex)
            {
                log.WriteToCVS(parent, x, ex.Message, "ERROR", LogManager.GetFileOwner(x));               
            }
        }
        public static void ProcessFileInfo(string parent, string x, MigrationSettigs d, LogCsvFile log)
        {
            try
            {
                string i = "blank";
                string end_part = x.Substring(x.LastIndexOf("\\"));
                if (end_part.LastIndexOf(".") != -1)
                    i = end_part.Substring(end_part.LastIndexOf("."));

                if (x.Length < 260)
                {
                    FileInfo f = new FileInfo(x);
                    if (d.DoIncremental)
                    {
                        if (f.LastWriteTime < DateTime.Parse(d.LastDate))
                        {
                            log.WriteWarning("[excluded]: " + x);
                            return;
                        }
                    }
                    log.WriteToCVS(parent, x, i, f.Length.ToString(), LogManager.GetFileOwner(x), f.LastWriteTime.ToShortDateString());
                }
                else
                {
                    log.WriteWarning("[invalid_length]: " + x);
                }
            }
            catch (Exception ex)
            {
                log.WriteError("[" + x + "]:" + ex.Message);
            }
        }       
        public static void ProcessFileChanges(string parent, string x, MigrationSettigs d, LogCsvFile log)
        {
            if (d.DoIncremental)
            {
                DateTime from_date = DateTime.Parse(d.LastDate);
                FileInfo file = new FileInfo(x);
                if (file.LastWriteTime >= from_date)
                    log.WriteToCVS(parent, x, "object changed");
            }
        }
        public static void ProcessDeleteXlsxFiles(string parent, string x, MigrationSettigs d, LogCsvFile log)
        {
            if (x.Length < 260)
            {
                FileInfo file = new FileInfo(x);
                if (file.Extension=="xlsx")
                    file.Delete();

                log.WriteInfo("[deleting]: " + x);
                log.WriteToCVS(parent, x, "deleted");
            }
            else
            {
                if (x.EndsWith(".xlsx"))
                    File.Delete(x);

                log.WriteWarning("[invalid_length]: " + x);
                log.WriteToCVS(parent, x, "invalid url");
            }
        }

        public static void CheckInvalidFolders(string masterfile_path, string doTransit, string report_name)
        {
            MigrationSettigs d = Init(doTransit, "");
            LogCsvFile log = new LogCsvFile("Folder, Name", report_name);
            string[] excludedFolders = d.ExcludeFolders.Split(',');

            string[] lines = System.IO.File.ReadAllLines(masterfile_path);
            foreach (string idpair in lines)
            {
                string[] values = idpair.Split(new char[] { ',' });
                string id = values[0];

                try
                {                    
                    string folderPath = string.Format(d.SearchPattern, id);
                    string parent = folderPath.Substring(folderPath.LastIndexOf("\\") + 1);
                    foreach (string folder in Directory.GetDirectories(folderPath, "*", SearchOption.AllDirectories))
                    {
                        DirectoryInfo dir = new DirectoryInfo(folder);
                        var found = from f in excludedFolders
                                    where f.Equals(dir.Name, StringComparison.CurrentCultureIgnoreCase)
                                    select f;

                        if(found.Count()>0)
                            log.WriteToCVS(parent, folder);
                    }
                   
                }
                catch (Exception ex)
                {                   
                    log.WriteError(id);
                    log.WriteError(ex.Message);
                }
            }
        }

        #endregion


        #region File Size

        public static void GetTransitsFolderSize(string textfilewithids, string report_path)
        {
            MigrationSettigs d = Init("1", "");
            LogCsvFile log = new LogCsvFile("parent, size", report_path);
            GetDirecotrySizeFromTextFileList(textfilewithids, d, log);
        }
        public static void GetUsersFolderSize(string textfilewithids, string report_path)
        {
            MigrationSettigs d = Init("0","");
            LogCsvFile log = new LogCsvFile("parent, size", report_path);
            GetDirecotrySizeFromTextFileList(textfilewithids, d, log);
        }
        static void GetDirecotrySizeFromTextFileList(string textfilewithids, MigrationSettigs d, LogCsvFile log)
        {
            long count = 1;
            log.WriteInfo("[opening file]");
            string[] lines = System.IO.File.ReadAllLines(textfilewithids);
            log.WriteInfo(string.Format("[{0} records found]", lines.Length.ToString()));
            string folder = string.Empty;
        
            foreach (string idpair in lines)
            {
                try
                {
                    string[] values = idpair.Split(new char[] { ',' });
                    string id = values[0];

                    log.WriteInfo(string.Format("[processing {0}/{1}]", count.ToString(), lines.Length.ToString()));
                    folder = string.Format(d.SearchPattern, id);
                    long size = DirSize(folder, true, d, log);
                    log.WriteToCVS(folder, size.ToString());
                    log.WriteInfo("[done]");
                }

                catch (Exception ex)
                {
                    log.WriteToCVS(folder, "error");
                    log.WriteError("[" + folder + "]: " + ex.Message);
                }
                finally
                {
                    count++;
                }
            }
                
            
        }
        private static long DirSize(string sourceDir, bool recurse, MigrationSettigs d, LogCsvFile log)
        {
            long size = 0;
            string[] fileEntries = Directory.GetFiles(sourceDir);

            foreach (string fileName in fileEntries)
            {
                if (fileName.Length < 256)
                {                    
                    Interlocked.Add(ref size, (new FileInfo(fileName)).Length);                   
                }
                else
                {
                    log.WriteWarning("[invalid length] :" + fileName);
                }
            }

            if (recurse)
            {
                string[] subdirEntries = Directory.GetDirectories(sourceDir);

                Parallel.For<long>(0, subdirEntries.Length, () => 0, (i, loop, subtotal) =>
                {
                    if (subdirEntries[i].Length < 256)
                    {
                        if ((File.GetAttributes(subdirEntries[i]) & FileAttributes.ReparsePoint) != FileAttributes.ReparsePoint)
                        {
                            subtotal += DirSize(subdirEntries[i], true, d, log);
                            return subtotal;
                        }
                    }
                    else
                    {
                        log.WriteWarning("[invalid length] :" + subdirEntries[i]);
                    }
                    return 0;
                },
                    (x) => Interlocked.Add(ref size, x)
                );
            }
            return size;
        }

        #endregion


        #region All Files

        public static void GetMissingDirectories(string masterfile_path, string doTransit)
        {
            MigrationSettigs d = Init(doTransit, "");
            LogCsvFile log = new LogCsvFile("Directory, Exists", "missing_directories");          
            string[] lines = System.IO.File.ReadAllLines(masterfile_path);
           
            foreach (string id in lines)
            {
                try
                {
                    string folderPath = string.Format(d.SearchPattern, id);
                    if (Directory.Exists(folderPath))
                    {
                        log.WriteToCVS(folderPath, "TRUE");
                    }
                    else
                    {
                        log.WriteToCVS(string.Format(d.SearchPattern, id), "FALSE");
                    }
                }
                catch (Exception ex)
                {
                    log.WriteToCVS(string.Format(d.SearchPattern, id), "ERROR");
                    log.WriteError(id);
                    log.WriteError(ex.Message);                   
                }
            }
        }
        public static void GetFiles(string masterfile_path, string doTransit, string from_date, ProcessFileDelegate FileHandler, string report_name)
        {
            MigrationSettigs d = Init(doTransit, from_date);
            LogCsvFile log = new LogCsvFile("", report_name);
            string[] lines = System.IO.File.ReadAllLines(masterfile_path);           

            foreach (string idpair in lines)
            {
                string[] values = idpair.Split(new char[] { ',' });
                string id = values[0];

                string folderPath = string.Format(d.SearchPattern, id);
                string parentFolderName = folderPath.Substring(folderPath.LastIndexOf("\\") + 1);
               
                log.WriteInfo("processing " + parentFolderName + "...");
                GetFiles(folderPath, parentFolderName, d, log, FileHandler);
                log.WriteInfo("done...");
            }
         
        }
        static void GetFiles(string sourceDir, string parent, MigrationSettigs d, LogCsvFile log, ProcessFileDelegate FileHandler)
        {            
            string[] fileEntries = TryGetDirFiles(sourceDir, log);
            foreach (string x in fileEntries)
                FileHandler(parent, x, d, log);                                 

            string[] subdirEntries = TryGetSubDirs(sourceDir, log);
            foreach (string subdir in subdirEntries)
                GetFiles(subdir, parent, d, log, FileHandler);
        }
      
        static string[] TryGetDirFiles(string sourceDir, LogCsvFile log)
        {
            try
            {
                return ( Directory.GetFiles(sourceDir) );
            }
            catch (Exception ex)
            {
                log.WriteError(sourceDir);
                log.WriteError(ex.Message);

                return new string[] { };
            }
        }      
        static string[] TryGetSubDirs(string sourceDir, LogCsvFile log)
        {
            try
            {
                return (Directory.GetDirectories(sourceDir));
            }
            catch (Exception ex)
            {
                log.WriteError(sourceDir);
                log.WriteError(ex.Message);

                return new string[] { };
            }
        }


        #endregion

       
    }

    #region helper classes

    class MigrationSettigs
    {
        public string SearchPattern { get; set; }    
        public string ExcludeFolders { get; set; }              
        public string JobsSourcePsPathKeyValue { get; set; }
        public string JobsSourcePsDisplayUrlKeyValue { get; set; }
        public string JobsSourcePsUrlKeyValue { get; set; }
        public string JobsTargetPsPathKeyValue { get; set; }
        public string JobsTargetPsDisplayUrlKeyValue { get; set; }
        public string JobsTargetPsUrlKeyValue { get; set; }
        public string LastDate { get; set; }       

        public bool DoIncremental
        {
            get
            {
                if (!string.IsNullOrEmpty(LastDate))
                {                    
                    return true;
                }
                return false;
            }
        }
    }
   
    #endregion

}
