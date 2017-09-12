using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MarketWatchAuto
{
    class Program
    {
        static void Main(string[] args)
        {
            string tag = "https://finance.google.com/finance/info?q=NSE:";
            ReadTheSheet(tag);
        }

        private static Dictionary<string, Dictionary<string, string>> _formattedWebData;
        private static List<Dictionary<string, Dictionary<string, string>>> _listForWebData;
        private static Dictionary<string, int> _result;
        private static Dictionary<string, List<string>> historicalList;

        public static void CheckAndCreateDirectories(string path, string text, out string fullPath)
        {
            string filename = path + text;
            fullPath = filename;

            Directory.CreateDirectory(filename);
        }

        private static void DoProcess(Dictionary<string, List<string>> historyList)
        {
            //var historicalList = new Dictionary<string, List<string>>();
            //var closeValueList = new List<string>();
            //string failedYahooSymbolLog = @"C:\MarketWatch\GoogleResult\Log\YahooHistoryLog";
            //string failedYahooSymbolLogText = "\\FailedYahooSymbolLog.txt";
            //string fullPath = string.Empty;
            //var suffixNSE = ".NS";
            //string value;
            //int i = 1;
            //var tag = HistoricalStockDownloader.CompleteChartTag();
            //CheckAndCreateDirectories(failedYahooSymbolLog, String.Empty, out fullPath);

            //using (StreamWriter sw = File.CreateText(fullPath + failedYahooSymbolLogText))
            //{
            //    foreach (var data in _listForWebData)
            //    {
            //        foreach (var validData in data)
            //        {
            //            try
            //            {
            //                var historicalData = HistoricalStockDownloader.DownloadData(validData.Key + suffixNSE, tag);
            //                if (historicalData == null)
            //                {
            //                    continue;
            //                }
            //                closeValueList.AddRange(historicalData.Select(stock => stock.Close));
            //                validData.Value.TryGetValue("CurrentPrice", out value);
            //                closeValueList.Insert(0, value);
            //                if (!historicalList.ContainsKey(validData.Key))
            //                {
            //                    historicalList.Add(validData.Key, closeValueList);
            //                }
            //                closeValueList = new List<string>();
            //            }
            //            catch (Exception)
            //            {
            //                sw.WriteLine(validData.Key);
            //            }
            //        }
            //    }
            //}
            CompareForPercIncreaseForLongDur(historicalList);
        }

        //private static void CompareForPercIncrease(Dictionary<string, List<string>> historicalList)
        //{
        //    bool firstInterv = false, secndInterv = false, thirdInterv = false, fourthInterv = false;
        //    var boolStatus = new Dictionary<string, List<bool>>();
        //    _result = new Dictionary<string, int>();
        //    foreach (var hist in historicalList)
        //    {
        //        var histValue = hist.Value;
        //        int i = 0;
        //        if (histValue.Count > 1 && float.Parse(histValue[i]) >= float.Parse(histValue[i + 1]))
        //            firstInterv = true;
        //        if (histValue.Count > 2 && float.Parse(histValue[i + 1]) >= float.Parse(histValue[i + 2]))
        //            secndInterv = true;
        //        if (histValue.Count > 3 && float.Parse(histValue[i + 2]) >= float.Parse(histValue[i + 3]))
        //            thirdInterv = true;
        //        if (histValue.Count > 4 && float.Parse(histValue[i + 3]) >= float.Parse(histValue[i + 4]))
        //            fourthInterv = true;

        //        boolStatus.Add(hist.Key, new List<bool>() { firstInterv, secndInterv, thirdInterv, fourthInterv });
        //        firstInterv = false;
        //        secndInterv = false;
        //        thirdInterv = false;
        //        fourthInterv = false;
        //    }
        //    _result = CalculatePriority(boolStatus);
        //    //var created = CreatePriorityExcel(_result);
        //    //if (created)
        //    //{
        //    //    Console.WriteLine("The priority excel sheet can not be created. Hence process is incomplete.");
        //    //}
        //    //else
        //    //{
        //    //    Console.WriteLine("The process is completed successfully.");
        //    //}
        //}

        private static void CompareForPercIncreaseForLongDur(Dictionary<string, List<string>> historicalList)
        {
            bool firstInterv = false, secndInterv = false, thirdInterv = false, fourthInterv = false,
                fifthInterv = false, sixthInterv = false, seventhInterv = false, eightInterv = false,
                ninethInterv = false, tenthInterv = false, firstMonthInterv = false, secMonthInterv = false,
                thirdMonthInterv = false, fourMonthInterv = false, fiveMonthInterv = false, sixMonthInterv = false,
                sevenMonthInterv = false; 
            var boolStatus = new Dictionary<string, List<bool>>();
            int monthInterv = 30;
            _result = new Dictionary<string, int>();
            foreach (var hist in historicalList)
            {
                var histValue = hist.Value;
                int i = 0;
                if (histValue.Count > 1 && float.Parse(histValue[i]) >= float.Parse(histValue[i + 1]))
                    firstInterv = true;
                if (histValue.Count > 2 && float.Parse(histValue[i + 1]) >= float.Parse(histValue[i + 2]))
                    secndInterv = true;
                if (histValue.Count > 3 && float.Parse(histValue[i + 2]) >= float.Parse(histValue[i + 3]))
                    thirdInterv = true;
                if (histValue.Count > 4 && float.Parse(histValue[i + 3]) >= float.Parse(histValue[i + 4]))
                    fourthInterv = true;
                if (histValue.Count > 5 && float.Parse(histValue[i + 4]) >= float.Parse(histValue[i + 5]))
                    fifthInterv = true;
                if (histValue.Count > 6 && float.Parse(histValue[i + 5]) >= float.Parse(histValue[i + 6]))
                    sixthInterv = true;
                if (histValue.Count > 7 && float.Parse(histValue[i + 6]) >= float.Parse(histValue[i + 7]))
                    seventhInterv = true;
                if (histValue.Count > 8 && float.Parse(histValue[i + 7]) >= float.Parse(histValue[i + 8]))
                    eightInterv = true;
                if (histValue.Count > 9 && float.Parse(histValue[i + 8]) >= float.Parse(histValue[i + 9]))
                    ninethInterv = true;
                if (histValue.Count > 10 && float.Parse(histValue[i + 9]) >= float.Parse(histValue[i + 10]))
                    tenthInterv = true;

                if (histValue.Count > 30 && float.Parse(histValue[i + 29]) >= float.Parse(histValue[i + 30]))
                    firstMonthInterv = true;
                if (histValue.Count > monthInterv * 2 && float.Parse(histValue[i + (monthInterv*2)-1]) >= float.Parse(histValue[i + (monthInterv * 2)]))
                    secMonthInterv = true;
                if (histValue.Count > monthInterv * 3 && float.Parse(histValue[i + (monthInterv * 3) - 1]) >= float.Parse(histValue[i + (monthInterv * 3)]))
                    thirdMonthInterv = true;
                if (histValue.Count > monthInterv * 4 && float.Parse(histValue[i + (monthInterv * 4) - 1]) >= float.Parse(histValue[i + (monthInterv * 4)]))
                    fourMonthInterv = true;
                if (histValue.Count > monthInterv * 6 && float.Parse(histValue[i + (monthInterv * 6) - 1]) >= float.Parse(histValue[i + (monthInterv * 6)]))
                    fiveMonthInterv = true;
                if (histValue.Count > monthInterv * 8 && float.Parse(histValue[i + (monthInterv * 8) - 1]) >= float.Parse(histValue[i + (monthInterv * 8)]))
                    sixMonthInterv = true;
                if (histValue.Count > monthInterv * 10 && float.Parse(histValue[i + (monthInterv * 10) - 1]) >= float.Parse(histValue[i + (monthInterv * 10)]))
                    sevenMonthInterv = true;

                boolStatus.Add(hist.Key, new List<bool>() { firstInterv, secndInterv, thirdInterv, fourthInterv, fifthInterv, sixthInterv,
                    seventhInterv, eightInterv, ninethInterv, tenthInterv,firstMonthInterv,secMonthInterv,thirdMonthInterv,fourMonthInterv,fiveMonthInterv,sixMonthInterv,sevenMonthInterv });
                firstInterv = false;
                secndInterv = false;
                thirdInterv = false;
                fourthInterv = false;
                fifthInterv = false;
                sixthInterv = false;
                seventhInterv = false;
                eightInterv = false;
                ninethInterv = false;
                tenthInterv = false;
                firstMonthInterv = false;
                secMonthInterv = false;
                thirdMonthInterv = false;
                fourMonthInterv = false;
                fiveMonthInterv = false;
                sixMonthInterv = false;
                sevenMonthInterv = false;
            }
            _result = CalculatePriorityForLongDur(boolStatus);
            //var created = CreatePriorityExcel(_result);
            //if (created)
            //{
            //    Console.WriteLine("The priority excel sheet can not be created. Hence process is incomplete.");
            //}
            //else
            //{
            //    Console.WriteLine("The process is completed successfully.");
            //}
        }

        //private static Dictionary<string, int> CalculatePriority(Dictionary<string, List<bool>> boolStatus)
        //{   
        //    var result = new Dictionary<string, int>();
        //    foreach (var boolStatu in boolStatus)
        //    {
        //        int i = 0;
        //        int prio = 0;
        //        var histValue = boolStatu.Value;
        //        if (histValue[i] && histValue[i + 1] && histValue[i + 2] && histValue[i + 3])
        //        {
        //            prio = 4;
        //        }
        //        else if ((histValue[i] && histValue[i + 1] && histValue[i + 2]) ||
        //                 (histValue[i + 1] && histValue[i + 2] && histValue[i + 3]))
        //        {
        //            prio = 3;
        //        }
        //        else if (histValue[i] && histValue[i + 1])
        //        {
        //            prio = 2;
        //        }
        //        else if (histValue[i + 1] && histValue[i + 2])
        //        {
        //            prio = 1;
        //        }
        //        else
        //        {
        //            prio = 0;
        //        }
        //        result.Add(boolStatu.Key, prio);
        //    }
        //    return result;
        //}

        private static Dictionary<string, int> CalculatePriorityForLongDur(Dictionary<string, List<bool>> boolStatus)
        {
            var result = new Dictionary<string, int>();
            foreach (var boolStatu in boolStatus)
            {
                int i = 0;
                int prio = 0;
                var histValue = boolStatu.Value;
                var longResult = (histValue[10] && histValue[i + 11] && histValue[i + 12] && histValue[i + 13] && histValue[i + 14] && histValue[i + 15] && histValue[i + 16]);

                if (histValue[i] && histValue[i + 1] && histValue[i + 2] && histValue[i + 3] && histValue[i + 4] && histValue[i + 5] && histValue[i + 6] && histValue[i + 7]
                    && histValue[i + 8] && histValue[i + 9] && longResult)
                {
                    prio = 11;
                }
                else if (histValue[i] && histValue[i + 1] && histValue[i + 2] && histValue[i + 3] && histValue[i + 4] && histValue[i + 5] && histValue[i + 6] && histValue[i + 7] 
                    && histValue[i + 8] && histValue[i + 9])
                {
                    prio = 10;
                }
                else if ((histValue[i] && histValue[i + 1] && histValue[i + 2] && histValue[i + 3] && histValue[i + 4] && histValue[i + 5] && histValue[i + 6] && histValue[i + 7]
                    && histValue[i + 8])|| 
                    (histValue[i + 1] && histValue[i + 2] && histValue[i + 3] && histValue[i + 4] && histValue[i + 5] && histValue[i + 6] && histValue[i + 7]
                    && histValue[i + 8] && histValue[i + 9]))
                {
                    prio = 9;
                }
                else if ((histValue[i] && histValue[i + 1] && histValue[i + 2] && histValue[i + 3] && histValue[i + 4] && histValue[i + 5] && histValue[i + 6] && histValue[i + 7]) 
                    ||(histValue[i + 2] && histValue[i + 3] && histValue[i + 4] && histValue[i + 5] && histValue[i + 6] && histValue[i + 7]
                    && histValue[i + 8] && histValue[i + 9]))
                {
                    prio = 8;
                }
                else if (histValue[i] && histValue[i + 1] && histValue[i + 2] && histValue[i + 3] && histValue[i + 4] && histValue[i + 5] && histValue[i + 6])
                {
                    prio = 7;
                }
                else if (histValue[i] && histValue[i + 1] && histValue[i + 2] && histValue[i + 3] && histValue[i + 4] && histValue[i + 5])
                {
                    prio = 6;
                }
                else if (histValue[i] && histValue[i + 1] && histValue[i + 2] && histValue[i + 3] && histValue[i + 4])
                {
                    prio = 5;
                }
                else if (histValue[i] && histValue[i + 1] && histValue[i + 2] && histValue[i + 3] && histValue[i + 4])
                {
                    prio = 4;
                }
                else if ((histValue[i] && histValue[i + 1] && histValue[i + 2]) ||
                         (histValue[i + 1] && histValue[i + 2] && histValue[i + 3]))
                {
                    prio = 3;
                }
                else if (histValue[i] && histValue[i + 1])
                {
                    prio = 2;
                }
                else if (histValue[i + 1] && histValue[i + 2])
                {
                    prio = 1;
                }
                else
                {
                    prio = 0;
                }
                result.Add(boolStatu.Key, prio);
            }
            return result;
        }

        private static bool CreatePriorityExcel(Dictionary<string, int> result)
        {
            var isCreated = true;
            int j = 0;
            var listOfNewShetColumns = new List<string>() { "Symbol", "Priority" };
            string priorityExcelPath = @"C:\MarketWatch\GoogleResult";
            string priorityExcelName = "\\PriorityExcel.xlsx";
            string fullPath = string.Empty;
            Excel.Worksheet xlNewSheet = null;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = null;

            CheckAndCreateDirectories(priorityExcelPath, String.Empty, out fullPath);
            xlWorkBook = xlApp.Workbooks.Open(fullPath + priorityExcelName, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true,
                false, 0, true, false, false);
            Excel.Sheets workSheets = xlWorkBook.Worksheets;

            try
            {
                Console.WriteLine("Creating priority excel sheet");
                xlApp.DisplayAlerts = false;
                xlNewSheet = (Excel.Worksheet)workSheets.Add(workSheets[1], Type.Missing, Type.Missing, Type.Missing);
                xlNewSheet.Move(Missing.Value, xlWorkBook.Sheets[xlWorkBook.Sheets.Count]);
                xlNewSheet.Name = DateTime.Now.ToString("s").Replace('-', '.').Replace(':', '.');
                xlNewSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item[xlWorkBook.Sheets.Count];
                xlNewSheet.Select();

                // For header Names.

                for (int i = 0; i < listOfNewShetColumns.Count; i++)
                {
                    xlNewSheet.Cells[1, ++j] = listOfNewShetColumns[i];
                }

                xlNewSheet.Columns.AutoFit();

                // For Header colors
                j = 0;
                for (int i = 0; i < listOfNewShetColumns.Count; i++)
                {
                    xlNewSheet.Cells[1, ++j].Interior.Color = Excel.XlRgbColor.rgbBlanchedAlmond;
                }

                int k = 2;
                foreach (var key in result.Keys)
                {
                    var i1 = k;
                    Console.WriteLine("Updating the priority excel sheet column1: " + i1 + " ...");
                    xlNewSheet.Cells[k, 1] = key;
                    k++;
                }

                k = 2;
                foreach (var value in result.Values)
                {
                    var i1 = k;
                    Console.WriteLine("Updating the priority excel sheet column2: " + i1 + " ...");
                    xlNewSheet.Cells[k, 2] = value;
                    k++;
                }
            }
            catch (Exception e)
            {
                isCreated = false;
                Console.WriteLine("Exception: " + e.Message);
            }

            finally
            {
                xlWorkBook.Close();
                releaseObject(xlNewSheet);
                releaseObject(workSheets);
                releaseObject(xlWorkBook);
            }
            return isCreated;
        }

        public static bool ReadTheSheet(string tag)
        {
            Console.WriteLine("Initializing...");

            Stopwatch watch = Stopwatch.StartNew();
            Excel.Application xlApp;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range range;

            _listForWebData = new List<Dictionary<string, Dictionary<string, string>>>();
            historicalList = new Dictionary<string, List<string>>();
            bool isOperationCompleted = false;
            string[] str = new string[1764];
            var results = new List<string>();
            var strDetail = new Dictionary<string, List<string>>();

            string symbolExcelPath = @"C:\MarketWatch\GoogleResult";
            string symbolExcel = "\\BSE_Equity_List.xlsx";
            string failedSymbolLog = @"C:\MarketWatch\GoogleResult\Log";
            string failedSymbolLogText = "\\FailedSymbolLog.txt";
            string fullPath = string.Empty;
            string filePath = string.Empty;

            var listOfNewSheetColumns = new List<string>()
            {
                "ID",
                "Symbol",
                "CurrentPrice",
                "Date",
                "Time",
                "Change",
                "ChangePercentage(%)",
                "PreviousClosurePrice",
                "SecurityName",
                "Group",
                "Industry"
            };
            var symbolListsFinal = new List<string>();
            WebClient webClient = new WebClient() { Proxy = WebRequest.GetSystemWebProxy() };
            webClient.Proxy.Credentials = CredentialCache.DefaultCredentials;
            xlApp = new Excel.Application();

            CheckAndCreateDirectories(symbolExcelPath, String.Empty, out fullPath);
            filePath = fullPath + symbolExcel;
            Console.WriteLine("Opening File...");

            try
            {
                xlWorkbook = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t",
                    false, false, 0, true, 1, 0);
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.Item[1];
                range = xlWorksheet.UsedRange;

                Parallel.For(2, range.Rows.Count,
                    cCnt => { str[cCnt] = (string)(range.Cells[cCnt, 2] as Excel.Range).Value2; });

                Console.WriteLine("File Opened...");

                for (int i = 2; i <= range.Rows.Count; i++)
                {
                    var i1 = i;
                    Console.WriteLine("Updating the detail equity list for: " + str[i1] + " ...");

                    if (!strDetail.ContainsKey((string)(range.Cells[i1, 2] as Excel.Range).Value2))
                        strDetail.Add(
                            (string)(range.Cells[i1, 2] as Excel.Range).Value2,
                            new List<string>()
                            {
                                (string) (range.Cells[i1, 3] as Excel.Range).Value2,
                                (string) (range.Cells[i1, 5] as Excel.Range).Value2,
                                (string) (range.Cells[i1, 8] as Excel.Range).Value2
                            });
                }

                CheckAndCreateDirectories(failedSymbolLog, String.Empty, out fullPath);
                using (StreamWriter sw = File.CreateText(fullPath + failedSymbolLogText))
                {
                    for (int i = 2; i < str.Length; i++)
                    {
                        if (!string.IsNullOrEmpty(str[i]))
                        {
                            var i1 = i;
                            Console.WriteLine("Downloading equity request for: " + str[i1] + " ...");

                            try
                            {
                                string result;
                                if (str[i].Contains('&'))
                                {
                                    var splitSymbol = str[i].Split('&');
                                    var appendSymbolWithSpecialChar = splitSymbol[0] + "%26" + splitSymbol[1];
                                    result = Extract(tag + appendSymbolWithSpecialChar, webClient);
                                    results.Add(result);
                                    historicalList.Add(result, new List<string>());
                                }
                                else
                                {
                                    result = Extract(tag + str[i], webClient);
                                    results.Add(result);
                                }
                            }
                            catch (Exception)
                            {
                                sw.WriteLine(str[i]);
                            }
                        }
                    }
                }

                Console.WriteLine("Creating symbol data list for priority calculation...");

                for (int i = xlWorkbook.Worksheets.Count; i >= 1; i--)
                {
                    xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.Item[i];
                    range = xlWorksheet.UsedRange;
                    string key = string.Empty;
                    int j = 0;
                    foreach (var cCnt in historicalList)
                    {
                        key= (string)(range.Cells[cCnt, 2] as Excel.Range).Value2;
                        if (historicalList.ContainsKey(key))
                        {
                            historicalList[key][j++] = (string)(range.Cells[cCnt, 2] as Excel.Range).Value2;
                        }
                    }
                }
                
                Console.WriteLine("Creating symbol data list...");

                foreach (var result in results)
                {
                    string symbolLists;
                    _formattedWebData = GetWebDatas(result, out symbolLists);
                    _listForWebData.Add(_formattedWebData);
                    symbolListsFinal.Add(symbolLists);
                }

                if (_listForWebData.Any())
                {
                    Console.WriteLine("Creating result excel page...");

                    var creationResult = CreateNewSheet(strDetail, xlApp, listOfNewSheetColumns, _listForWebData,
                        symbolListsFinal, filePath, "1", historicalList)
                        ? "Operation Successfull!"
                        : "Operation Un-Successful!";
                    watch.Stop();
                    Console.WriteLine(creationResult + "...");
                    Console.WriteLine(creationResult + '\n' + "Time elapsed: " + watch.Elapsed);
                    isOperationCompleted = true;
                }
                else
                {
                    Console.WriteLine("No result has been performed!!");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
            finally
            {
                xlApp.Quit();
                releaseObject(xlWorksheet);
                releaseObject(xlWorkbook);
                releaseObject(xlApp);
            }

            return isOperationCompleted;
        }

        private static bool CreateNewSheet(Dictionary<string, List<string>> strDetail, Microsoft.Office.Interop.Excel.Application xlApp,
            List<string> listOfNewSheetColumns, List<Dictionary<string, Dictionary<string, string>>> listForWebData,
            List<string> symbolListsFinal, string filePath, string dialogResult, Dictionary<string, List<string>> historyList)
        {
            var isCreated = true;
            int j = 0;
            Excel.Worksheet xlNewsheet = null;
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false,
                Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Sheets workSheets = xlWorkbook.Worksheets;
            try
            {
                xlApp.DisplayAlerts = false;
                xlNewsheet = (Excel.Worksheet)workSheets.Add(workSheets[1], Type.Missing, Type.Missing, Type.Missing);
                xlNewsheet.Move(Missing.Value, xlWorkbook.Sheets[xlWorkbook.Sheets.Count]);
                xlNewsheet.Name = DateTime.Now.ToString("s").Replace('-', '.').Replace(':', '.');
                xlNewsheet = (Excel.Worksheet)xlWorkbook.Worksheets.Item[xlWorkbook.Sheets.Count];
                xlNewsheet.Select();

                // For header Names.

                for (int i = 0; i < listOfNewSheetColumns.Count; i++)
                {
                    xlNewsheet.Cells[1, ++j] = listOfNewSheetColumns[i];
                }

                xlNewsheet.Columns.AutoFit();

                // For Header colors
                j = 0;
                for (int i = 0; i < listOfNewSheetColumns.Count; i++)
                {
                    xlNewsheet.Cells[1, ++j].Interior.Color = Excel.XlRgbColor.rgbBlanchedAlmond;
                }

                //writing the web data respective to their column.
                int k = 0;
                for (int i = 2; i < listForWebData.Count + 2; i++)
                {
                    var i1 = i;
                    Console.WriteLine("Writing teh web data respective to their column for row: " + i1 + "...");
                    int m = 0, l = 0;
                    var symbolDictionary = listForWebData[k][symbolListsFinal[k]];
                    for (j = 1; j <= listOfNewSheetColumns.Count; j++)
                    {
                        if (j <= listOfNewSheetColumns.Count - 3)
                        {
                            xlNewsheet.Cells[i, j] = symbolDictionary[listOfNewSheetColumns[m]];
                        }
                        else
                        {
                            xlNewsheet.Cells[i, j] = strDetail.ContainsKey(symbolListsFinal[k])
                                    ? strDetail[symbolListsFinal[k]][l]
                                    : string.Empty;
                            l++;
                        }
                        m++;
                    }
                    k++;
                }
                //if (dialogResult == "0") //need to check for the Yahoo working api.
                //{
                    var lastRow = listOfNewSheetColumns.Count + 1;
                    xlNewsheet.Cells[1, lastRow] = "Priority";
                    xlNewsheet.Columns.AutoFit();

                    //For Header colors
                    xlNewsheet.Cells[1, lastRow].Interior.Color = Excel.XlRgbColor.rgbBlanchedAlmond;
                    DoProcess(historyList);
                    //writing the priority data respective to their column.
                    k = 2;
                    foreach (var value in _result.Values)
                    {
                        var i1 = k;
                        Console.WriteLine("Updating the priority excel sheet column: " + i1 + " ...");
                        xlNewsheet.Cells[k, lastRow] = value;
                        k++;
                    }
                //}
                xlNewsheet.Columns.AutoFit();
                xlWorkbook.Save();

            }
            catch (Exception e)
            {
                isCreated = false;
                Console.WriteLine("Exception: " + e.Message);
            }
            finally
            {
                xlWorkbook.Close();
                releaseObject(xlNewsheet);
                releaseObject(workSheets);
                releaseObject(xlWorkbook);
            }
            return isCreated;
        }

        private static string Extract(string googleHttpRequestString, WebClient webClient)
        {
            var strm = webClient.DownloadString(new Uri(googleHttpRequestString));
            return strm;
        }

        private static Dictionary<string, Dictionary<string, string>> GetWebDatas(string input, out string symbolList)
        {
            string symbolListToBeAssign = String.Empty;
            symbolList = null;
            var resultSymbolDictionary = new Dictionary<string, Dictionary<string, string>>();
            var resultDictionary = new Dictionary<string, string>();
            var replaceLine = Regex.Replace(input, "\n", String.Empty);
            var aaa = Regex.Replace(replaceLine, "\"", String.Empty);
            var ab = aaa.Replace("// [{", String.Empty).Replace("}]", String.Empty);

            var aa = ab.Split(',');
            string symbol = String.Empty;
            for (int i = 0; i < aa.Length; i++)
            {
                var str = aa[i];
                if (!str.Contains(":"))
                {
                    continue;
                }
                var strSplit = str.Split(':');
                var firstValue = strSplit[0].Trim();
                var value = strSplit[1].Trim();
                switch (firstValue)
                {
                    case "id":
                        resultDictionary.Add("ID", value);
                        break;
                    case "t":
                        value = value.Contains("\\u00026") ? value.Replace("\\u0026", "&") : value;
                        symbol = value;
                        resultDictionary.Add("Symbol", value);
                        symbolListToBeAssign = value;
                        break;
                    case "e":
                        resultDictionary.Add("Sensex", value);
                        break;
                    case "l_fix":
                        resultDictionary.Add("CurrentPrice", value);
                        break;
                    case "lt":
                        var time = aa[++i].Trim();
                        resultDictionary.Add("Date", value);
                        resultDictionary.Add("Time", time);
                        break;
                    case "c":
                        resultDictionary.Add("Change", value);
                        break;
                    case "cp":
                        resultDictionary.Add("ChangePercentage(%)", value);
                        break;
                    case "pcls_fix":
                        resultDictionary.Add("PreviousClosurePrice", value);
                        break;

                }

            }
            resultSymbolDictionary.Add(symbol, resultDictionary);
            symbolList = symbolListToBeAssign;
            return resultSymbolDictionary;
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception e)
            {
                obj = null;
                Console.WriteLine("Unable to release the Object " + e.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }

    public class HistoricalStock
    {
        public string Date { get; set; }
        public string Open { get; set; }
        public string High { get; set; }
        public string Low { get; set; }
        public string Close { get; set; }
        public string Volume { get; set; }
        public string AdjClose { get; set; }
    }

    public class HistoricalStockDownloader
    {
        public static List<HistoricalStock> DownloadData(string ticker, string tag)
        {
            var retval = new List<HistoricalStock>();
            //string failedYahooSymbolLog = @"C:\MarketWatch\Google\YahooHistoryLog";
            //string failedYahooSymbolLogText = "\\FailedYahooSymbolLog.txt";
            //string fullPath = string.Empty;

            //Business.CheckAndCreateDirectories(failedYahooSymbolLog, String.Empty, out fullPath);

            using (WebClient web = new WebClient())
            {
                //using (StreamWriter sw = File.CreateText(fullPath+ failedYahooSymbolLogText))
                //{
                //try
                //{
                ticker = ticker.Contains("\\u0026") ? ticker.Replace("\\u0026", "&") : ticker;
                string data = web.DownloadString(string.Format(tag, ticker));
                data = data.Replace("r", "");
                string[] rows = data.Split('\n');

                //First row is headers so ignore it.
                for (int i = 1; i < rows.Length; i++)
                {
                    if (rows[i].Replace("\n", "").Trim() == "") continue;
                    {
                        string[] cols = rows[i].Split(',');
                        var hs = new HistoricalStock()
                        {
                            Date = cols[0],
                            Open = cols[1],
                            High = cols[2],
                            Low = cols[3],
                            Close = cols[4],
                            Volume = cols[5],
                            AdjClose = cols[6]
                        };
                        retval.Add(hs);
                    }
                }
                return retval;
                //}
                //catch (Exception)
                //{
                //    sw.WriteLine(ticker);
                //}
                //return null;
                //}
            }
        }

        public static string CompleteChartTag()
        {
            int initialDayConst = 5, daysInMonth = 30;
            // in order to adjust the yahoo attributes.
            int negateForYahooattr = 1;
            string[] currentDay = new string[] { DateTime.Today.Day.ToString(), (DateTime.Today.Month - negateForYahooattr).ToString(), DateTime.Today.Year.ToString() };
            string[] initialDay = new string[]
            {
                DateTime.Today.Day<initialDayConst?(DateTime.Today.Day+daysInMonth-initialDayConst).ToString():(DateTime.Today.Day-initialDayConst).ToString(),
                DateTime.Today.Day<initialDayConst?(DateTime.Today.Month-initialDayConst-negateForYahooattr).ToString():(DateTime.Today.Month-negateForYahooattr).ToString(),
                DateTime.Today.Year.ToString()
            };
            var firstTag = "http://ichart.finance.yahoo.com/table.csv?g=d";
            string[] initialDayTag = new string[] { "&a=" + initialDay[1], "&b=" + initialDay[0], "&c=" + initialDay[2] };
            string[] currentDayTags = new string[] { "&d=" + currentDay[1], "&e=" + currentDay[0], "&f=" + currentDay[2] };

            var attributeTag = initialDayTag[0] + initialDayTag[1] + initialDayTag[2] + currentDayTags[0] +
                               currentDayTags[1] + currentDayTags[2];
            var lastTag = "&s={0}";
            var finalTag = firstTag + attributeTag + lastTag;
            return finalTag;
        }
    }
}
