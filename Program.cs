using static System.Net.Mime.MediaTypeNames;
using System.Runtime.CompilerServices;
using Syncfusion.Licensing;
using Syncfusion.XlsIO;

namespace SpreadsheetsFileUnion
{
    internal class Program
    {
        private const ConsoleColor ErrorForeColor = ConsoleColor.Red;
        private const ConsoleColor SuccessForeColor = ConsoleColor.Green;
        private static void Main(string[] args)
        {
            bool isDebug = true;
            string currentfolder = Environment.CurrentDirectory;
            var allFolder = Directory.EnumerateDirectories(currentfolder);
            if (isDebug)
            {
                Console.WriteLine("<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
                Console.WriteLine("                   'Merge Tool' file csv,xlsx                                               ");
                Console.WriteLine("<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
                Console.WriteLine("");
                Console.WriteLine("Start scan current directory");
            }
            List<string> discardedCSVFile = new List<string>();
            List<string> discardedxlsxFile = new List<string>();
            foreach (string folder in allFolder)
            {
                if (isDebug)
                {
                    Console.WriteLine("<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
                    Console.WriteLine("Found folder : " + Path.GetFileName(folder));
                }
                try
                {
                    var allCsv = from fname in Directory.EnumerateFiles(folder, "*.csv", SearchOption.AllDirectories)
                                 where !fname.Contains(Path.GetFileName(folder) + "_")
                                 select fname;
                    if (allCsv.Any())
                    {
                        if (isDebug)
                        {
                            Console.WriteLine("Found file CSV");
                        }
                        string masterHeader2 = File.ReadLines(allCsv.First()).First((string l) => !string.IsNullOrWhiteSpace(l));
                        if (isDebug)
                        {
                            Console.WriteLine("Check header CSV file");
                        }
                        IProgressBar progressBar2 = new ConsoleProgressBar(allCsv.Count());
                        for (int k = 0; k < allCsv.Count(); k++)
                        {
                            string f = allCsv.ElementAt(k);
                            progressBar2.ShowProgress(k + 1);
                            string currentHead2 = File.ReadLines(f).First((string l) => !string.IsNullOrWhiteSpace(l));
                            if (masterHeader2 != currentHead2)
                            {
                                discardedCSVFile.Add(f);
                            }
                        }
                        Console.WriteLine("");
                        string[] header = new string[1] { File.ReadLines(allCsv.First()).First((string l) => !string.IsNullOrWhiteSpace(l)) };
                        allCsv = from fname in Directory.EnumerateFiles(folder, "*.csv", SearchOption.AllDirectories)
                                 where !fname.Contains(Path.GetFileName(folder) + "_") && discardedCSVFile.IndexOf(fname) == -1
                                 select fname;
                        Console.WriteLine("Start merge csv");
                        var mergedData = allCsv.SelectMany((string csv) => File.ReadLines(csv).SkipWhile((string l) => string.IsNullOrWhiteSpace(l)).Skip(1));
                        var mergedFileContent = header.Concat(mergedData);
                        string[] s = mergedFileContent.ToArray();
                        string unionFileCsvName = $"{currentfolder}\\{Path.GetFileName(folder)}\\{Path.GetFileName(folder)}_mergedfile_{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv";

                        File.WriteAllLines(unionFileCsvName, s);
                        if (isDebug)
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("File merge completed successfully. File created : ");
                            Console.WriteLine(Path.GetFileName(unionFileCsvName) ?? "");
                            Console.ForegroundColor = ConsoleColor.White;
                        }
                    }
                    else if (isDebug)
                    {
                        Console.WriteLine("Nothing csv file in this directory");
                    }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ErrorForeColor;
                    Console.WriteLine("Error : " + ex.ToString());
                    Console.ForegroundColor = ConsoleColor.White;
                }

                try
                {
                    var allxls = from fname in Directory.EnumerateFiles(folder, "*.xlsx", SearchOption.AllDirectories)
                                 where !fname.Contains(Path.GetFileName(folder) + "_")
                                 select fname;
                    if (allxls.Any())
                    {
                        if (isDebug)
                            Console.WriteLine("Files XLS found");

                        using (ExcelEngine excelEngine = new ExcelEngine())
                        {
                            IApplication application = excelEngine.Excel;
                            application.DefaultVersion = ExcelVersion.Excel2016;

                            #region Check Header XLSX
                            if (isDebug)
                            {
                                Console.WriteLine("Check header file XLS");
                            }
                            string[] masterHeader = GetExcelHeader(allxls.First(), application);
                            IProgressBar progressBar = new ConsoleProgressBar(allxls.Count());
                            for (int j = 0; j < allxls.Count(); j++)
                            {
                                progressBar.ShowProgress(j + 1);
                                string file = allxls.ElementAt(j);
                                string[] currentHead = GetExcelHeader(file, application);
                                if (!masterHeader.SequenceEqual(currentHead))
                                {
                                    discardedxlsxFile.Add(file);
                                }
                            }
                            #endregion

                            #region Merge XLSX
                            Console.WriteLine("");
                            if (isDebug)
                            {
                                Console.WriteLine("Start merge XLSX file");
                            }
                            IWorkbook workbook = application.Workbooks.Create(1);
                            IWorksheet worksheet = workbook.Worksheets[0];
                            int lastRow = 1;
                            bool firstrow = true;
                            progressBar = new ConsoleProgressBar(allxls.Count());
                            for (int i = 0; i < allxls.Count(); i++)
                            {
                                progressBar.ShowProgress(i + 1);
                                string file2 = allxls.ElementAt(i);
                                if (discardedxlsxFile.IndexOf(file2) == -1)
                                {
                                    FileStream fileStream = new FileStream(file2, FileMode.Open, FileAccess.Read);
                                    IWorkbook workbooktemp = application.Workbooks.Open(fileStream, ExcelOpenType.Automatic);
                                    fileStream.Dispose();
                                    int startrow = (firstrow ? 1 : 2);
                                    if (workbooktemp.Worksheets[0].AutoFilters.FilterRange != null)
                                    {
                                        workbooktemp.Worksheets[0].AutoFilters.FilterRange = null;
                                    }
                                    IRange range = workbooktemp.Worksheets[0].Range[startrow, 1, workbooktemp.Worksheets[0].UsedRange.LastRow, workbooktemp.Worksheets[0].UsedRange.LastColumn];
                                    range.CopyTo(worksheet.Range[lastRow, 1]);
                                    lastRow = ((!firstrow) ? (lastRow + (workbooktemp.Worksheets[0].UsedRange.LastRow - 1)) : (lastRow + workbooktemp.Worksheets[0].UsedRange.LastRow));
                                    firstrow = false;
                                }
                            }
                            string unionFileXlsxName = $"{currentfolder}\\{Path.GetFileName(folder)}\\{Path.GetFileName(folder)}_mergedfile_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";
                            FileStream outputStream = new FileStream(unionFileXlsxName, FileMode.Create, FileAccess.ReadWrite);
                            workbook.SaveAs(outputStream);
                            workbook.Close();
                            outputStream.Dispose();
                            if (isDebug)
                            {
                                Console.WriteLine("");
                                Console.ForegroundColor = SuccessForeColor;
                                Console.WriteLine("File merge compleato con successo. File creato:");
                                Console.WriteLine(Path.GetFileName(unionFileXlsxName) ?? "");
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                            #endregion
                        }

                    }
                    else if (isDebug)
                    {
                        Console.WriteLine("Nessun file XLS");
                    }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ErrorForeColor;
                    Console.WriteLine("Errore eleborazione file XLS: " + ex.ToString());
                    Console.ForegroundColor = ConsoleColor.White;
                }                
            }
            if (discardedCSVFile.Count > 0 || discardedxlsxFile.Count > 0)
            {
                Console.WriteLine("<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
                Console.ForegroundColor = ErrorForeColor;
                Console.WriteLine("Attenzione! I seguenti file con header differente sono stati scartati : ");
                Console.WriteLine(string.Join("\r\n", discardedxlsxFile) ?? "");
                Console.WriteLine(string.Join("\r\n", discardedCSVFile) ?? "");
                Console.ForegroundColor = ConsoleColor.White;
            }
            Console.WriteLine("<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
            Console.WriteLine("                                   Premi ENTER per uscire                                   ");
            Console.WriteLine("<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
            Console.ReadLine();
        }
        private static string[] GetExcelHeader(string filepath, IApplication application)
        {
            FileStream fileStream = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = application.Workbooks.Open(fileStream, ExcelOpenType.Automatic);
            IWorksheet worksheet = workbook.Worksheets[0];
            int rowCount = worksheet.UsedRange.Rows.Length;
            int columnCount = worksheet.UsedRange.Columns.Length;
            string[] masterHeader = new string[columnCount];
            for (int i = 0; i < columnCount; i++)
            {
                masterHeader[i] = worksheet.Range[1, i + 1].Value;
            }
            return masterHeader;
        }
    }
}