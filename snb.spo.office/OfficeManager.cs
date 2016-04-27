using System;
using System.IO;
using System.Linq;
using System.Configuration;
using ExcelObj = Microsoft.Office.Interop.Excel;
using bns.spo.diagnostic;
using bns.spo.files;


namespace bns.spo.office
{
    class OfficeManager
    {
            
        public static void ConvertXlsToXlsxFromFileList(string textfilewithids, string doTransit, string from_lastdate, string report_path)
        {
            ExcelObj.Application excelApp = null;
            LogCsvFile log = new LogCsvFile("parent, file, comments", report_path);
            MigrationSettigs d = FileManager.Init(doTransit, from_lastdate);

            try
            {                                
                string[] lines = System.IO.File.ReadAllLines(textfilewithids);
                string startedAt = DateTime.Now.ToString();
                long count = 1;

                log.WriteInfo("[opening master file]");
                log.WriteInfo(string.Format("[{0} records found]", lines.Length.ToString()));

                excelApp = new ExcelObj.Application();
                excelApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                excelApp.AskToUpdateLinks = false;
                excelApp.DisplayAlerts = false;
                excelApp.Visible = false;
                

                foreach (string idpair in lines)
                {
                    string[] values = idpair.Split(new char[] { ',' });
                    string id = values[0];

                    string folder = string.Format(d.SearchPattern, id);
                    string parentFolder = folder.Substring(folder.LastIndexOf("\\") + 1);

                    log.WriteInfo("[processing]: " + parentFolder);
                    log.WriteInfo(string.Format("[processing]: {0}/{1}", count.ToString(), lines.Length.ToString()));

                    NewExcelInstance(excelApp, parentFolder, folder, 0, d, log);
                    count++;
                }

                log.WriteInfo("[Started at]: " + startedAt);
                log.WriteInfo("[Finished at]: " + DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                log.WriteError(ex.Message);
            }
            finally
            {
                try
                {
                    excelApp.Quit();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch(Exception e)
                {
                    log.WriteError(e.Message);
                }
            }
        }
        public static void CheckXlsToXlsxConversion(string masterfile_path, string doTransit, string from_lastdate)
        {
            MigrationSettigs d = FileManager.Init(doTransit, from_lastdate);
            LogCsvFile log = new LogCsvFile("parent, file, comments", "checkxlstoxlsx");
            string[] lines = System.IO.File.ReadAllLines(masterfile_path);

            log.WriteInfo("[Started at]: " + DateTime.Now.ToString());
            log.WriteInfo("[Parent Folders Found]: " + lines.Count().ToString());

            int convertedCount = 0;
            foreach (string idpair in lines)
            {
                string[] values = idpair.Split(new char[] { ',' });
                string id = values[0];

                string folderPath = string.Format(d.SearchPattern, id);
                string parentFolderName = folderPath.Substring(folderPath.LastIndexOf("\\") + 1);
                var files = Directory.EnumerateFiles(folderPath, "*xls", SearchOption.AllDirectories);

                log.WriteInfo("[processing]: " + parentFolderName + "...");

                int fileCount = 0;
                foreach (string x in files)
                {
                    try
                    {
                        if (x.Substring(x.LastIndexOf(".")).ToLower() == ".xls")
                        {
                            string converted = "false";
                            string xlsxFilePath = x.Replace(".xls", ".xlsx");
                            string xlsmFilePath = x.Replace(".xls", ".xlsm");

                            if (File.Exists(xlsxFilePath) || File.Exists(xlsmFilePath))
                            {
                                converted = "true";
                                convertedCount++;
                            }

                            fileCount++;

                            log.WriteToCVS(parentFolderName, x, converted);
                            log.WriteInfo("[" + x + "]" + " ...done ");
                        }
                    }
                    catch (Exception ex)
                    {
                        log.WriteInfo(x);
                        log.WriteInfo(ex.ToString());

                    }

                }

            }

            log.WriteInfo("[xlsx files found]: " + convertedCount.ToString());
            log.WriteInfo("[Finished at]: " + DateTime.Now.ToString());
        }
        public static void ListXlsPwd(string textfilewithids, string doTransit, string from_lastdate, string report_path)
        {
            ExcelObj.Application excelApp = null;
            LogCsvFile log = new LogCsvFile("parent,file,owner", report_path);
            MigrationSettigs d = FileManager.Init(doTransit, from_lastdate);

            try
            {
                string[] lines = System.IO.File.ReadAllLines(textfilewithids);
                string startedAt = DateTime.Now.ToString();

                //excelApp = new ExcelObj.Application();
                //excelApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                //excelApp.AskToUpdateLinks = false;
                //excelApp.DisplayAlerts = false;
                //excelApp.Visible = false;

                foreach (string idpair in lines)
                {
                    string[] values = idpair.Split(new char[] { ',' });
                    string id = values[0];

                    string folder = string.Format(d.SearchPattern, id);
                    string parentFolder = folder.Substring(folder.LastIndexOf("\\") + 1);

                    NewExcelInstance(excelApp, parentFolder, folder, 1, d, log);
                }

                log.WriteInfo("[Started]: " + startedAt);
                log.WriteInfo("[Finished]: " + DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                log.WriteError(ex.Message);
            }
            finally
            {
                try
                {
                    //excelApp.Quit();
                    //GC.Collect();
                    //GC.WaitForPendingFinalizers();
                }
                catch (Exception e)
                {
                    log.WriteError(e.Message);
                }
            }
        }



        private static string ConvertFile(ExcelObj.Application excelApp, string parentFolderPath, string inputFile, MigrationSettigs d, LogCsvFile log)
        {            
            var xlsxNewFile = string.Empty;
            string fname = inputFile.Substring(inputFile.LastIndexOf("\\") + 1);
            string status = Constants.NotProcessed;
            FileInfo i = new FileInfo(inputFile);
            ExcelObj.Workbook excelWrkBk = null;

            try
            {
                Console.WriteLine("processing: " + fname);
                if (FileHasBeenConverted(i) &&
                    string.IsNullOrEmpty(d.LastDate))
                {
                    return Constants.AlreadyConverted;
                }

                excelWrkBk = excelApp.Workbooks.Open(
                   inputFile,
                   0,
                   Type.Missing,
                   Type.Missing,
                   "1234567890",
                   Type.Missing,
                   Type.Missing,
                   Type.Missing,
                   Type.Missing,
                   Type.Missing,
                   Type.Missing,
                   Type.Missing,
                   Type.Missing,
                   Type.Missing,
                   Type.Missing);


                ExcelObj.XlFileFormat format =
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook;


                if (excelWrkBk.HasVBProject)
                {
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
                    xlsxNewFile = inputFile.Replace(".xls", ".xlsm");
                }
                else if(i.Extension ==".xls")
                {
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook;
                    xlsxNewFile = inputFile.Replace(".xls", ".xlsx");
                }
                else if (i.Extension == ".xlt")
                {
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook;
                    xlsxNewFile = inputFile.Replace(".xlt", "");
                    xlsxNewFile = xlsxNewFile + "_template.xlsx";
                }               

                excelWrkBk.SaveAs(xlsxNewFile,
                    format,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    ExcelObj.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing);

                status = Constants.Converted;

                log.WriteToCVS(parentFolderPath, inputFile, "converted");
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2146827284)
                {
                    log.WriteInfo("password protected file found...");
                    log.WriteToCVS(parentFolderPath, inputFile, "password protected");
                    status = Constants.PasswordFound;
                }
                else
                {
                    log.WriteToCVS(parentFolderPath, inputFile, "error");
                    status = Constants.NotProcessed;
                }
              
                log.WriteError(inputFile);
                log.WriteError(ex.Message);
            }
            finally
            {
                try
                {
                    Console.WriteLine("closing: " + fname);
                    if (excelWrkBk != null)
                        excelWrkBk.Close(false, Type.Missing, Type.Missing);
                }
                catch (Exception e)
                {
                    log.WriteError(inputFile);
                    log.WriteError(e.Message);
                }
              
            }

            return status;
        }      
        private static void NewExcelInstance(ExcelObj.Application excelApp, string parentFolderPath, string path, int operation, MigrationSettigs d, LogCsvFile log)
        {                        
            long count = 0;
            long filesWithPwdCount = 0;
            long filesConvertedCount = 0;
            long fileAlreadyConverted = 0;
            long filesNotProcessed = 0;

            bool doIncremental = false;
            DateTime from_date = DateTime.MinValue;

            try
            {
                log.WriteInfo("[processing]: " + parentFolderPath + " (" + DateTime.Now.ToString() + ")");

                var files = Directory.EnumerateFiles(path, "*.xl*", SearchOption.AllDirectories).Where(x => x.Length < 260);
                if (d.DoIncremental) from_date = DateTime.Parse(d.LastDate);

                string status = Constants.NotProcessed;
                foreach (string x in files)
                {
                    if (x.Substring(x.LastIndexOf(".")).ToLower() == ".xls" 
                        || x.Substring(x.LastIndexOf(".")).ToLower() == ".xlt")
                    {
                        if (doIncremental)
                        {

                            FileInfo file = new FileInfo(x);
                            if (file.LastWriteTime < from_date)
                                continue;
                        }

                        count++;
                        if (operation == 0)
                        {

                            status = ConvertFile(excelApp, parentFolderPath, x, d, log);
                        }
                        else if (operation == 1)
                        {

                            status = CheckFileForPassword(excelApp, parentFolderPath, x, d, log);
                        }

                        switch (status)
                        {
                            case Constants.NotProcessed:                              
                                filesNotProcessed++;
                                break;
                            case Constants.Converted:
                                filesConvertedCount++;
                                break;
                            case Constants.PasswordFound:
                                filesWithPwdCount++;
                                break;
                            case Constants.AlreadyConverted:
                                fileAlreadyConverted++;
                                break;
                            default:
                                filesNotProcessed++;
                                break;
                        }
                    }

                }

                log.WriteInfo("[finished]: " + parentFolderPath + " (" + DateTime.Now.ToString() + ")");
            }
            catch (Exception ex)
            {
                log.WriteError(ex.Message);
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                log.WriteInfo("[Total xls files found]: " + count.ToString());
                log.WriteInfo("[Total files converted]: " + filesConvertedCount.ToString());
                log.WriteInfo("[Total files with passwords]: " + filesWithPwdCount.ToString());
                log.WriteInfo("[Total files already converted]: " + fileAlreadyConverted.ToString());

                if (operation == 0)
                {

                    log.WriteInfo("[Total files with errors]: " + filesNotProcessed.ToString());
                }
                else
                {

                    log.WriteInfo("[Total files w/o passwords]: " + filesNotProcessed.ToString());
                }
            }
                
           
        }
        private static bool FileHasBeenConverted(FileInfo x)
        {
            //overrides check
            if (ConfigurationManager.AppSettings["conversion.checkifexists"]=="false")
                return false;

            string filePath = string.Empty;
            if (x.Extension == ".xls")
            {
                return (File.Exists(x.FullName.Replace(".xls", ".xlsx")) || File.Exists(x.FullName.Replace(".xls", ".xlsm")));
            }
            else if (x.Extension == ".xlt")
            {
                return (File.Exists(x.FullName.Replace(".xlt", ".xlsx")));
            }
            else if (x.Extension == ".xltx")
            {
                return (File.Exists(x.FullName.Replace(".xltx", ".xlsx")));
            }
            else
                return false;

        }
                    
        private static string CheckFileForPassword(ExcelObj.Application excelApp, string parentFolderPath, string inputFile, MigrationSettigs d, LogCsvFile log)
        {
            var missing = System.Reflection.Missing.Value;
            string fname = inputFile.Substring(inputFile.LastIndexOf("\\") + 1);
            string status = Constants.NotProcessed;
            ExcelObj.Workbook excelWrkBk = null;

            try
            {

                //Console.WriteLine("[processing]: " + fname);               
                //excelWrkBk = excelApp.Workbooks.Open(
                //    inputFile,
                //    0,
                //    true,
                //    missing,
                //    "124567890",
                //    missing,
                //    missing,
                //    missing,
                //    missing,
                //    missing,
                //    missing,
                //    missing,
                //    missing,
                //    missing,
                //    missing);


                if (OfficePasswordHelper.IsPasswordProtected(inputFile))
                {
                    Console.WriteLine("password protected file found...");
                    log.WriteToCVS(parentFolderPath, inputFile, LogManager.GetFileOwner(inputFile));
                    status = Constants.PasswordFound;
                }
                else
                {

                    status = Constants.NotProcessed;
                }
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2146827284)
                {
                    Console.WriteLine("password protected file found...");
                    log.WriteToCVS(parentFolderPath, inputFile, LogManager.GetFileOwner(inputFile));
                    status = Constants.PasswordFound;
                }
                else
                {                    
                    log.WriteError(inputFile);
                    log.WriteError(ex.Message);
                }                            
            }
            finally
            {
                try
                {
                    if (excelWrkBk != null)
                        excelWrkBk.Close(false, missing, missing);
                }
                catch (Exception e)
                {
                    log.WriteError(inputFile);
                    log.WriteError(e.Message);
                }
            }

            return status;
        }


    }
}
