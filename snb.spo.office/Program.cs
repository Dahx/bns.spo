using System;
using System.Collections.Generic;
using bns.spo.office;
using bns.spo.files;
using bns.spo.migration;


namespace bns.spo
{
    class Program
    {

        static void Main(string[] args)
        {
            try
            {               
                string o;
                List<string> p = GetOperationAndParameters(args, out o);
                switch (o.ToLower())
                {
                    case "-xlstoxlsx":
                        OfficeManager.ConvertXlsToXlsxFromFileList(p[0], p[1], p[2], p[3]);
                        break;

                    case "-delete.xlsx":
                        FileManager.ProcessFileDelegate process_file_deletions = FileManager.ProcessDeleteXlsxFiles;                      
                        FileManager.GetFiles(p[0], p[1], p[2], process_file_deletions, "deleted_files");                      
                        Console.WriteLine("done.");
                        break;

                    case "-checkxlsconversion":
                        OfficeManager.CheckXlsToXlsxConversion(p[0], p[1], p[2]);
                        break;

                    case "-listpwd":
                        OfficeManager.ListXlsPwd(p[0],p[1],p[2], p[3]);
                        break;                                 
                
                    case "-getmissingdirs":
                        FileManager.GetMissingDirectories(p[0],p[1]);
                        break;
                
                    case "-check":
                        FileManager.ProcessFileDelegate process_file_validations = FileManager.ProcessFileValidations;
                        Console.WriteLine("checking validations...");
                        FileManager.GetFiles(p[0], p[1], p[2], process_file_validations, p[3]);
                        FileManager.CheckInvalidFolders(p[0], p[1], p[3]);
                        Console.WriteLine("done.");                                                    
                        break;                   

                    case "-getfiles":
                        FileManager.ProcessFileDelegate process_fileInfo = FileManager.ProcessFileInfo;
                        FileManager.GetFiles(p[0], p[1], p[2], process_fileInfo, p[3]);
                        break;

                    case "-job.gen":
                        JobManager.GeneratePsJobs(p[0], p[1], p[2], p[3]);
                        break;

                    case "-job.run":
                        JobManager j = new JobManager(p[0]);
                        j.ProcessJobs();
                        break;

                    case "-users.getsize":
                        FileManager.GetUsersFolderSize(p[0], p[1]);
                        break;                   

                    case "-transits.getsize":
                        FileManager.GetTransitsFolderSize(p[0], p[1]);
                        break;

                    default:
                        Console.Write("Invalid operation...");                        
                        break;

                }

                Console.WriteLine("Press enter to continue...");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadLine();
            }
        }

  
        #region helper methods

        
        /// <summary>
        /// Helper to wirte color text to the console
        /// </summary>
        /// <param name="message"></param>
        /// <param name="color"></param>
        /// <param name="addline"></param>
        static void ConsoleWriteColor(string message, ConsoleColor color, bool addline = false)
        {
            var lastForegroundColor = Console.ForegroundColor;
            Console.ForegroundColor = color;

            if (!addline)
            {
                Console.Write(message);
            }
            else
            {
                Console.WriteLine(message);
            }

            Console.ForegroundColor = lastForegroundColor;
        }
        /// <summary>
        /// parses the entries and returns the operation and its parameters
        /// </summary>
        /// <param name="args"></param>
        /// <param name="operation"></param>
        /// <returns></returns>
        static List<string> GetOperationAndParameters(string[] args, out string operation)
        {
            List<string> parameters = new List<string>();
            operation = args[0];

            if (args.Length > 1)
            {
                for (int i = 1; i <= args.Length - 1; i++)
                {
                    parameters.Add(args[i]);
                }
            }

            return parameters;
        }
       

        #endregion


    }
}
