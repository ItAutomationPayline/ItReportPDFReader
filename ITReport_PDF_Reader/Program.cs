using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using OfficeOpenXml;
using Microsoft.Office.Interop.Excel;
using static UglyToad.PdfPig.Core.PdfSubpath;
using System.Text;

namespace ITReport_PDF_Reader
{
    public class Program
    {
        public static string basePath = AppDomain.CurrentDomain.BaseDirectory;
        public static string sourceFolder = basePath;
        public static string outputFilePath = Path.Combine(basePath, "Output");
        public static string filePath1 = "";
        
        public static void Main(string[] args)
        {
            Console.WriteLine("After pasting PDF file, Press Enter to start processing.");
            Console.ReadLine();

            var inputFile = Directory.GetFiles(Path.Combine(sourceFolder, "Input"), "*.pdf")
                                     .OrderByDescending(f => new FileInfo(f).CreationTime)
                                     .ToList();

            if (inputFile.Count == 0)
            {
                Console.WriteLine("Required file not found. Ensure there is .pdf file in Input folder.");
                return;
            }

            filePath1 = inputFile.First();
            Console.WriteLine($"Reading PDF: {Path.GetFileName(filePath1)}");

            // Process PDF and get results
            var records = ExtractDataFromPdf(filePath1);

            // Write to Excel
            WriteToExcel(records);

            Console.WriteLine("Processing completed. Excel file generated in Output folder.");
        }

        // 🔹 Extract PAN + Amounts from PDF
        public static List<(string Pan, string EmployeeId, string TaxAmount, string BaseAmount)> ExtractDataFromPdf(string pdfPath)
        {
            var results = new List<(string Pan, string EmployeeId, string TaxAmount, string BaseAmount)>();

            using (PdfDocument document = PdfDocument.Open(pdfPath))
            {
                var pages = document.GetPages().ToList();

                for (int i = 0; i < pages.Count; i++)
                {
                    string pageText = pages[i].Text;
                    string[] lines = pageText.Split('\n');
                    foreach (var line in lines)
                    {
                        if (line.Contains("Spot tax for the current processed month"))
                        {
                            // --- Extract Tax & Base amounts ---
                            var match = Regex.Match(line, @"Rs\.?\s?([\d,]+).*?Rs[:.]?\s?([\d,]+)");
                            if (match.Success)
                            {
                                string taxAmount = match.Groups[1].Value;
                                string baseAmount = match.Groups[2].Value;

                                string pan = "NOT FOUND";
                                string employeeId = "NOT FOUND";

                                // --- PAN Extraction (current + previous page) ---
                                string panSearchText = pageText + (i > 0 ? pages[i - 1].Text : "");
                                var panMatch = Regex.Match(panSearchText, @"PAN\s*:\s*([A-Z]{3}P[A-Z][0-9]{4}[A-Z])");
                                if (panMatch.Success)
                                {
                                    pan = panMatch.Groups[1].Value;
                                }
                                var searchText = new StringBuilder();
                                if (i > 0) searchText.AppendLine(pages[i - 1].Text ?? string.Empty);
                                searchText.AppendLine(pageText);
                                if (i < pages.Count - 1) searchText.AppendLine(pages[i + 1].Text ?? string.Empty);
                                string combined = searchText.ToString();
                                // --- Employee ID Extraction ---
                                // Look for digits (6–10 long, no commas) near "Below 60 years"
                                int idxYears = combined.IndexOf("years", StringComparison.OrdinalIgnoreCase);
                                int idxPan = -1;
                                if (idxYears >= 0)
                                {
                                    // search for PAN after idxYears; if not found, try any PAN
                                    idxPan = combined.IndexOf("PAN", idxYears, StringComparison.OrdinalIgnoreCase);
                                }
                                if (idxYears >= 0 && idxPan > idxYears)
                                {
                                    int start = idxYears + "years".Length;
                                    int len = idxPan - start;
                                    if (len > 0)
                                    {
                                        string between = combined.Substring(start, len);
                                         //extract first continuous digits sequence (no commas) from the between-substring
                                        var digitsMatch = Regex.Match(between, @"\d{4,12}");
                                        if (digitsMatch.Success)
                                        {
                                            employeeId = digitsMatch.Value;
                                        }
                                    }
                                }
                                else
                                {
                                    // Fallback: look for a digit-only line immediately after phrases like "Below 60 years"
                                    var fallback = Regex.Match(combined, @"Below\s*\d+\s*years[\s\S]{0,60}?(\b\d{4,12}\b)", RegexOptions.IgnoreCase);
                                    if (fallback.Success)
                                    {
                                        employeeId = fallback.Groups[1].Value;
                                    }
                                    else
                                    {
                                        // Another fallback: look for a standalone 6-10 digit number near where "PAN" appears
                                        // (in case 'years' isn't present exactly)
                                        int panIndexAny = combined.IndexOf("PAN", StringComparison.OrdinalIgnoreCase);
                                        if (panIndexAny > 0)
                                        {
                                            // take up to 100 chars before PAN and search for digits
                                            int startIdx = Math.Max(0, panIndexAny - 100);
                                            string beforePan = combined.Substring(startIdx, panIndexAny - startIdx);
                                            var dm = Regex.Match(beforePan, @"(\d{6,10})");
                                            if (dm.Success) employeeId = dm.Groups[1].Value;
                                        }
                                    }
                                }
                                // --- Add result if values are valid ---
                                if (Convert.ToDouble(taxAmount.Replace(",", "")) >  Convert.ToDouble(baseAmount.Replace(",", ""))*0.4)
                                {
                                    results.Add((pan, employeeId, taxAmount, baseAmount));
                                }
                            }
                        }
                    }
                }
            }

            return results;
        }
        // 🔹 Write extracted data to Excel
        public static void WriteToExcel(List<(string Pan,string empid, string TaxAmount, string BaseAmount)> records)
        {
            if (records.Count == 0)
            {
                Console.WriteLine("No matching records found in PDF.");
                return;
            }

            Directory.CreateDirectory(outputFilePath);

            string excelFile = Path.Combine(outputFilePath, "PdfReport"+ Path.GetFileName(filePath1.Replace(".pdf","")) + ".xlsx");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Report");

                    ws.Cells[1, 1].Value = "Employee ID";
                    ws.Cells[1, 2].Value = "PAN";
                    ws.Cells[1, 3].Value = "Tax Amount";
                    ws.Cells[1, 4].Value = "Base Amount";

                    int row = 2;
                    foreach (var rec in records)
                    {
                        ws.Cells[row, 1].Value = rec.empid;
                        ws.Cells[row, 2].Value = rec.Pan;
                        ws.Cells[row, 3].Value = Convert.ToDouble(rec.TaxAmount);
                        ws.Cells[row, 4].Value = Convert.ToDouble(rec.BaseAmount);
                        row++;
                    }

                    //ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    package.SaveAs(new FileInfo(excelFile));
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
