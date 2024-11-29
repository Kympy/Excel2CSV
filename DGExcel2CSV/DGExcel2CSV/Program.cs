using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace DGExcel2CSV
{
    internal class Program
    {
        private static Application app;
        private static Workbook workbook;

        private static string targetFolder;

        public static void Main(string[] args)
        {
            Console.WriteLine();
            Console.WriteLine("--------- Excel 2 CSV ---------");
            Console.WriteLine();

            if (args == null || args.Length == 0)
            {
                Console.WriteLine("No arguments");
                return;
            }

            if (args.Length > 2)
            {
                Console.WriteLine("Too many arguments.");
                return;
            }

            Console.WriteLine("Entered Arguments ---");
            for (int i = 0; i < args.Length; i++)
            {
                Console.WriteLine($"\t[{i + 1}] {args[i]}");
            }

            Application activeExcel = null;
            try
            {
                activeExcel = (Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                activeExcel = null;
            }
            if (activeExcel != null)
            {
                Console.WriteLine("** Excel process is running. **");
                return;
            }

            bool isFile = File.Exists(args[0]);
            bool isDirectory = Directory.Exists(args[0]);

            if (isFile == false && isDirectory == false)
            {
                Console.WriteLine($"Argument [0] is not a file or directory. {args[0]}");
                return;
            }

            // Output 경로가 주어지지 않았거나, null 일 경우
            if (args.Length == 1 || string.IsNullOrEmpty(args[1]))
            {
                // 파일의 상위폴더 혹은 직접 받은 폴더로 지정
                targetFolder = isFile ? Path.GetDirectoryName(args[0]) : args[0];
                // 새 폴더 생성
                targetFolder = Path.Combine(targetFolder, "_CSV");
                Console.WriteLine($"Make Temp CSV output folder : {targetFolder}");
            }
            else
            {
                targetFolder = args[1];
            }

            Console.WriteLine();
            Console.WriteLine($"Target folder : {targetFolder}");
            Console.WriteLine();

            if (Directory.Exists(targetFolder) == false)
            {
                Directory.CreateDirectory(targetFolder);
            }

            if (isFile)
            {
                Convert(args[0]);
            }
            else if (isDirectory)
            {
                var fileList = Directory.GetFiles(args[0], "*.xlsx");
                foreach (var file in fileList)
                {
                    if (file.IndexOf('$') != -1)
                        continue;
                    Convert(file);
                }
            }
            Console.WriteLine();
            Console.WriteLine("Finished.");
            Console.WriteLine();

            GC.Collect();

            return;
        }

        private static void Convert(string file)
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            workbook = app.Workbooks.Open(file);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];  // 첫 번째 시트

            // 사용된 범위만 가져오기
            var usedRange = worksheet.UsedRange;

            // 빈 행 및 열 제거
            TrimEmptyRowsAndColumns(usedRange);

            var csvPath = Path.ChangeExtension(file, ".csv");
            var csvFileName = Path.GetFileName(csvPath);
            var savePath = Path.Combine(targetFolder, csvFileName);
            savePath = savePath.Replace('/', '\\');
            Console.Write($"Converting... [{file}]");

            try
            {
                if (File.Exists(savePath))
                {
                    File.Delete(savePath);
                }

                var csvWriter = new StreamWriter(savePath, false, Encoding.UTF8);

                foreach (Microsoft.Office.Interop.Excel.Range row in usedRange.Rows)
                {
                    var rowValues = new List<string>();
                    foreach (Range cell in row.Cells)
                    {
                        var cellValue = cell.Text.ToString().Trim();
                        // Console.WriteLine(cellValue);
                        if (!string.IsNullOrEmpty(cellValue))  // 빈 셀은 추가하지 않음
                        {
                            rowValues.Add($"\"{cellValue}\"");
                        }
                        else
                        {
                            rowValues.Add("\"0\"");
                        }
                    }

                    if (rowValues.Count > 0)
                    {
                        csvWriter.WriteLine(string.Join(",", rowValues));  // 비어 있지 않은 행만 CSV에 저장
                    }
                    else
                    {
                        Console.WriteLine("Empty Table.");
                    }
                }

                csvWriter.Close();
                Console.WriteLine(" ...Saved!");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            workbook.Close(false);
            app.Quit();

            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(app);

            workbook = null;
            app = null;
        }

        private static void TrimEmptyRowsAndColumns(dynamic usedRange)
        {
            // 빈 행 및 열을 제거하는 로직
            for (int i = usedRange.Rows.Count; i >= 1; i--)
            {
                bool isEmptyRow = true;
                for (int j = 1; j <= usedRange.Columns.Count; j++)
                {
                    if (usedRange.Cells[i, j].Text.ToString().Trim() != "")
                    {
                        // Console.WriteLine(usedRange.Cells[i, j].Text.ToString());
                        isEmptyRow = false;
                        break;
                    }
                }

                if (isEmptyRow)
                {
                    // 빈 행 제거
                    usedRange.Rows[i].Delete();
                }
            }

            for (int j = usedRange.Columns.Count; j >= 1; j--)
            {
                bool isEmptyColumn = true;
                for (int i = 1; i <= usedRange.Rows.Count; i++)
                {
                    if (usedRange.Cells[i, j].Text.ToString().Trim() != "")
                    {
                        isEmptyColumn = false;
                        break;
                    }
                }

                if (isEmptyColumn)
                {
                    Console.WriteLine($"Empty Columns {j}");
                    // 빈 열 제거
                    usedRange.Columns[j].Delete();
                }
            }
        }
    }
}
