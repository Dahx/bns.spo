using System;
using System.IO;
using System.Text;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using bns.spo.diagnostic;
using System.Configuration;
using bns.spo.files;

namespace bns.spo.migration
{
    class JobManager
    {
      
        MigrationSettigs d;
        LogCsvFile log;
        string currentPs;
        string job_repository;
       

        public JobManager(string repository)
        {
            job_repository = repository;
            DirectoryInfo currentFolder = new DirectoryInfo(job_repository);

            d = FileManager.Init("1", "");
            log = new LogCsvFile("Machine,JobName,Status,Time,Date,Comments", Environment.MachineName);
            log.ReportPath = job_repository + "\\" + Environment.MachineName + ".csv";    
        }

        public static void GeneratePsJobs(string master_file, string job_template, string job_repository, string dotransit)
        {
            MigrationSettigs d = FileManager.Init(dotransit, "");
            LogCsvFile log = new LogCsvFile("", "job_conversion");
                     
            string[] lines = System.IO.File.ReadAllLines(master_file);
            foreach (string idpair in lines)
            {               
               
                log.WriteInfo("[processing]: " + idpair);
                StringBuilder ps_content = new StringBuilder();
                using (StreamReader ps = File.OpenText(job_template))
                {
                    ps_content.Append(ps.ReadToEnd());
                }

                string[] values = idpair.Split(new char[] { ',' });
                string id = values[0];
                string url = values[0];
                if (values.Length > 1) url = values[1];

                ps_content.Replace("##source_path##", string.Format(d.JobsSourcePsPathKeyValue, id));
                ps_content.Replace("##source_displayurl##", string.Format(d.JobsSourcePsDisplayUrlKeyValue, id));
                ps_content.Replace("##source_url##", string.Format(d.JobsSourcePsUrlKeyValue, id));

                ps_content.Replace("##target_path##", string.Format(d.JobsTargetPsPathKeyValue, url));
                ps_content.Replace("##target_displayurl##", string.Format(d.JobsTargetPsDisplayUrlKeyValue, url));
                ps_content.Replace("##target_url##", string.Format(d.JobsTargetPsUrlKeyValue, url));

                string ps_file = string.Format("{0}\\{1}.ps1", job_repository, id);
                using (FileStream f = File.Create(ps_file)) f.Close();
                File.AppendAllText(ps_file, ps_content.ToString());

            }
            log.WriteInfo("[done]");
        }

        public void ProcessJobs()
        {                      
            var jobs = Directory.GetFiles(job_repository, "*.ps1");
            foreach(string ps in jobs)                      
            {     
                currentPs = ps.Substring(ps.LastIndexOf("\\") + 1);                                            
                log.WriteInfo(string.Format("[initializing]: {0}", currentPs));
                log.WriteToCVS(Environment.MachineName, currentPs, "Initializing", DateTime.Now.ToLongTimeString(), DateTime.Today.ToShortDateString(),"");
                RunJobAsPs(ps);                                               
            }
                                           
        }
       
        private void RunJobAsPs(string jobpath)
        {
            Runspace runspace = null;
            try
            {
                runspace = RunspaceFactory.CreateRunspace();
                runspace.Open();
               
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.Runspace = runspace;
                    ps.AddScript(System.IO.File.ReadAllText(jobpath));
                    ps.InvocationStateChanged += Ps_InvocationStateChanged;

                    ps.Invoke();                            
                    WriteJobErrors(ps);
                }

            }
            catch (Exception ex)
            {
                log.WriteError(ex.Message);
            }
        }

        private void Ps_InvocationStateChanged(object sender, PSInvocationStateChangedEventArgs e)
        {           
            log.WriteInfo(e.InvocationStateInfo.State.ToString());
            string comments = string.Empty;            
            if (e.InvocationStateInfo.State == PSInvocationState.Completed)            
                comments = GetMetalogixJobLogResults(currentPs,"Failed");                           
            log.WriteToCVS(Environment.MachineName, currentPs, e.InvocationStateInfo.State.ToString(), DateTime.Now.ToLongTimeString(), DateTime.Today.ToShortDateString(), comments);
        }

        private string WriteJobErrors(PowerShell powershell)
        {
            try
            {
                var errors = powershell.Streams.Error.ReadAll();
                StringBuilder errorResults = new StringBuilder();

                if (errors.Count > 0)
                {
                    log.WriteToCVS(Environment.MachineName, currentPs, "Error", DateTime.Now.ToLongTimeString(), DateTime.Today.ToShortDateString());
                    foreach (var error in errors)
                    {
                        log.WriteError(error.ToString());
                        errorResults.AppendLine(error.ToString());
                    }
                }
                return errorResults.ToString();
            }
            catch (Exception ex)
            {
                log.WriteError(ex.ToString());  
            }

            return string.Empty;
        }
        
        private string WriteJobResults(System.Collections.ObjectModel.Collection<PSObject> results)
        {
            try
            {
                StringBuilder taskResults = new StringBuilder();
                if (results != null)
                {
                    foreach (var result in results)
                    {
                        log.WriteInfo(result.ToString());
                        taskResults.AppendLine(result.ToString());
                    }
                }
                return taskResults.ToString();
            }
            catch (Exception ex)
            {
                log.WriteError(ex.ToString());
            }
            return string.Empty;
        }

        private string GetMetalogixJobLogResults(string id, string status)
        {
            string jobresults = string.Empty;
            string errorFound = string.Empty;

            data.sqlce.Database mxDb = new data.sqlce.Database(ConfigurationManager.AppSettings["metalogix.log.connectionstring"]);
            string jobid = mxDb.GetValue("SELECT TOP 1 [JobID] FROM [Jobs] WHERE [Source] LIKE " + string.Format("'%{0}' AND [Status] = 'Done' ORDER BY [Finished] DESC", id));

            if (!string.IsNullOrEmpty(jobid))
            {
                jobresults = 
                    mxDb.GetValue(string.Format("SELECT [ResultsSummary] FROM [Jobs] WHERE [JobID] = '{0}'", jobid));
                errorFound = 
                    mxDb.GetValue(string.Format("SELECT COUNT(*) FROM [LogItems] WHERE [LogID] = '{0}' AND [Status] = '{1}'", jobid, status));
            }

            if (!string.IsNullOrEmpty(errorFound))
                return string.Format("{0} errors found. {1}",errorFound, jobresults);
            else
                return jobresults;
        }
       
        
    }

    
}
